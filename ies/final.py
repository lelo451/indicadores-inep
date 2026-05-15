# -*- coding: utf-8 -*-
"""Gera ``outputs/list_ies_final.xlsx`` combinando três fontes em ordem de
prioridade:

1. **e-MEC** (``Microdados/ies_siglas.csv``) — base oficial. Marca
   ``complemento='n'`` quando a IES tem ao menos um campo preenchido.
2. **Censo da Educação Superior** (``outputs/lista_ies_consolidada.xlsx``) —
   suplementa coluna a coluna as IES que e-MEC retornou ``not_found``.
   Marca ``complemento='y'`` se algum campo foi preenchido.
3. **Indicadores INEP** (``outputs/indicadores_consolidados.xlsx``) — para
   as IES que nem e-MEC nem Censo cobrem (descredenciadas, fundidas etc.),
   recupera Nome/Sigla/Org./Categ. das edições de ENADE/CPC/IDD/IGC em que
   a IES apareceu. Marca ``complemento='i'``.

Limpeza:
- Nomes do e-MEC frequentemente trazem anotações administrativas anexadas
  (ex.: "Em supervisão", "Suspensão contrato FIES", "Sub Judice"). O script
  trunca o nome a partir do primeiro desses marcadores. Ao final, imprime
  os nomes ainda mais longos que a média + 2·desvio padrão, para inspeção.
"""

from __future__ import annotations

import re

import pandas as pd

from normalizacao import (
    CATEG_ADMIN_CANONICAL,
    ORG_ACAD_CANONICAL,
    ORG_ACAD_LEGACY_CODES,
    map_canonical,
    strip_double_quotes,
    title_case_with_connectors,
)

from ._common import (
    BASE_CSV,
    CONSOLIDADA_XLSX,
    FINAL_XLSX,
    INDICADORES_XLSX,
    OUTPUTS_DIR,
)


OUTPUT_COLS = [
    "codigo_ies",
    "sigla",
    "nome",
    "organizacao",
    "categoria",
    "complemento",
]

BASE_RENAME = {
    "nome_emec": "nome",
    "organizacao_emec": "organizacao",
    "categoria_emec": "categoria",
}

SUPP_RENAME = {
    "Codigo da IES": "codigo_ies",
    "Sigla da IES": "sigla",
    "Nome da IES": "nome",
    "Organização Acadêmica": "organizacao",
    "Categoria Administrativa": "categoria",
}

# indicadores_consolidados.xlsx usa as mesmas colunas mas com "Código" acentuado.
INDIC_RENAME = {
    "Código da IES": "codigo_ies",
    "Sigla da IES": "sigla",
    "Nome da IES": "nome",
    "Organização Acadêmica": "organizacao",
    "Categoria Administrativa": "categoria",
}

INFO_COLS = ["sigla", "nome", "organizacao", "categoria"]

# Marcadores que indicam início de anotações administrativas anexadas ao nome.
# Tudo a partir do primeiro marcador encontrado é descartado.
NAME_TRAIL_MARKERS = [
    r"Em\s+supervis[ãa]o",
    r"Suspens[ãa]o\s+contrato\s+FIES",
    r"Suspens[ãa]o\s+PROUNI",
    r"Suspens[ãa]o\s+PRONATEC",
    r"Suspens[ãa]o\s+de\s+ingresso",
    r"Suspens[ãa]o\s+de\s+autonomia",
    r"Suspens[ãa]o\s+das?\s+prerrogativas",
    r"Credenciamento\s+EaD\s+Provis[óo]rio",
    r"Em\s+descredenciamento\s+volunt[áa]rio",
    r"Descredenciada",
    r"Acervo\s+Acad[êe]mico",
    r"Sub\s+Judice",
    r"Veda[çc][ãa]o\s+de",
]
_NAME_TRAIL_RE = re.compile("|".join(NAME_TRAIL_MARKERS), re.IGNORECASE)
# Prefixo "(NNN) " (e qualquer ruído antes dele) deve ser removido.
_CODE_PREFIX_RE = re.compile(r"^.*?\(\d+\)\s*", re.DOTALL)
_TRAIL_TRIM = " \t\n\r-–—,;:."


def _clean_nome(value: object) -> object:
    if not isinstance(value, str):
        return value
    match = _NAME_TRAIL_RE.search(value)
    if match:
        value = value[: match.start()]
    value = _CODE_PREFIX_RE.sub("", value, count=1)
    return value.strip(_TRAIL_TRIM)


def _report_long_names(df: pd.DataFrame) -> None:
    """Imprime nomes desproporcionalmente longos (possíveis sobras de
    informação não filtrada) para inspeção manual."""
    names = df["nome"].dropna()
    if names.empty:
        return
    lengths = names.str.len()
    mean, std = lengths.mean(), lengths.std()
    threshold = max(80, int(mean + 2 * std))
    suspeitos = df.loc[df["nome"].str.len() > threshold].copy() if std else df.iloc[0:0]
    suspeitos["__len"] = suspeitos["nome"].str.len()
    suspeitos = suspeitos.sort_values("__len", ascending=False).head(30)

    print()
    print(
        f"=== Nomes potencialmente longos demais "
        f"(threshold = {threshold} chars, média={mean:.1f}, std={std:.1f}) ==="
    )
    if suspeitos.empty:
        print("  Nenhum nome acima do limiar.")
        return
    print(f"  Mostrando {len(suspeitos)} nomes (mais longos primeiro):")
    for _, row in suspeitos.iterrows():
        print(f"  [{row['codigo_ies']:>6}] ({row['__len']:>3} chars) {row['nome']}")


def _load_base() -> pd.DataFrame:
    df = pd.read_csv(BASE_CSV)
    df = df.rename(columns=BASE_RENAME)
    return df[["codigo_ies", *INFO_COLS]].copy()


def _load_supplement() -> pd.DataFrame:
    df = pd.read_excel(CONSOLIDADA_XLSX, sheet_name="IES")
    df = df.rename(columns=SUPP_RENAME)
    df["codigo_ies"] = pd.to_numeric(df["codigo_ies"], errors="coerce").astype("Int64")
    return df[["codigo_ies", *INFO_COLS]].copy()


def _load_indicadores() -> pd.DataFrame:
    """Carrega indicadores_consolidados.xlsx reduzido a uma linha por IES.

    ``consolidar_indicadores.py`` já normaliza Nome/Sigla/Org./Categ. para o
    valor do ano mais recente, então basta agregar com ``first()`` para pegar
    o primeiro valor não-nulo por coluna em cada grupo de IES.

    Também aplica ORG_ACAD_LEGACY_CODES para sanar códigos numéricos antigos
    (ENADE 2004-2008) que escapam ao decodificador do consolidador.
    """
    df = pd.read_excel(INDICADORES_XLSX)
    df = df.rename(columns=INDIC_RENAME)
    df["codigo_ies"] = pd.to_numeric(df["codigo_ies"], errors="coerce").astype("Int64")
    df = df.dropna(subset=["codigo_ies"])
    df["organizacao"] = df["organizacao"].map(
        lambda v: ORG_ACAD_LEGACY_CODES.get(str(v).strip(), v) if pd.notna(v) else v
    )
    return df.groupby("codigo_ies", as_index=False)[INFO_COLS].first()


_INDIC_COL_MAP = {
    "sigla": ("Sigla da IES", strip_double_quotes),
    "nome": ("Nome da IES", title_case_with_connectors),
    "organizacao": (
        "Organização Acadêmica",
        lambda v: map_canonical(v, ORG_ACAD_CANONICAL),
    ),
    "categoria": (
        "Categoria Administrativa",
        lambda v: strip_double_quotes(map_canonical(v, CATEG_ADMIN_CANONICAL)),
    ),
}


def apply_to_indicadores(final_df: pd.DataFrame) -> None:
    """Reaplica nome/sigla/org./categ. de ``list_ies_final`` em
    ``indicadores_consolidados.xlsx``. Pula IES com ``complemento='i'``
    (informação veio dos próprios indicadores — seria circular). Valores são
    renormalizados após a cópia para preservar o estilo do arquivo destino."""
    if not INDICADORES_XLSX.exists():
        print(f"  {INDICADORES_XLSX.name} não encontrado — pulando.")
        return

    valid = final_df[final_df["complemento"] != "i"].copy()
    valid["codigo_ies"] = pd.to_numeric(valid["codigo_ies"], errors="coerce").astype("Int64")
    valid = valid.dropna(subset=["codigo_ies"])

    df_ind = pd.read_excel(INDICADORES_XLSX)
    cod = pd.to_numeric(df_ind["Código da IES"], errors="coerce").astype("Int64")

    # ORG_ACAD_LEGACY_CODES é aplicado independentemente do filtro de
    # complemento — códigos como "3"/"6" são genéricos e não dependem da
    # fonte. Caso contrário, IES com complemento='i' (ex.: 1985) ficariam
    # com o código numérico no indicadores enquanto list_ies_final já o
    # decodifica.
    if "Organização Acadêmica" in df_ind.columns:
        org_str = df_ind["Organização Acadêmica"].astype("string").str.strip()
        legacy_mask = org_str.isin(ORG_ACAD_LEGACY_CODES.keys())
        if legacy_mask.any():
            decoded = org_str[legacy_mask].map(ORG_ACAD_LEGACY_CODES).map(
                lambda v: map_canonical(v, ORG_ACAD_CANONICAL)
            )
            df_ind.loc[legacy_mask, "Organização Acadêmica"] = decoded
            print(
                f"  {int(legacy_mask.sum())} rótulo(s) Organização legados "
                f"decodificados via ORG_ACAD_LEGACY_CODES"
            )

    # Sigla=='0' aparece como sentinel de "não informado" em edições antigas.
    # Trata como ausente para que a lógica de fill abaixo o substitua quando
    # houver replacement, ou o converta em NaN quando não houver.
    sigla = df_ind["Sigla da IES"].astype("string").str.strip()
    sentinel_mask = sigla.isin(["0", "0.0"])
    if sentinel_mask.any():
        df_ind.loc[sentinel_mask, "Sigla da IES"] = pd.NA
        print(f"  {int(sentinel_mask.sum())} sigla(s) sentinel '0' convertidas em NaN")

    stats: dict[str, tuple[int, int]] = {}
    for src, (tgt, normalize) in _INDIC_COL_MAP.items():
        lookup = {
            int(k): normalize(v)
            for k, v in zip(valid["codigo_ies"], valid[src])
            if pd.notna(v) and (not isinstance(v, str) or v.strip())
        }
        lookup = {k: v for k, v in lookup.items() if pd.notna(v)}
        mapped = cod.map(lookup)
        cur = df_ind[tgt].astype("string").str.strip()
        cur_null = df_ind[tgt].isna() | cur.isin(["", "0", "0.0"])
        has_new = mapped.notna()
        filled = int((has_new & cur_null).sum())
        updated = int((has_new & ~cur_null & (cur != mapped.astype("string").str.strip())).sum())
        df_ind.loc[has_new, tgt] = mapped[has_new]
        stats[tgt] = (filled, updated)

    df_ind.to_excel(INDICADORES_XLSX, index=False)
    print(f"Reaplicado em {INDICADORES_XLSX.name}:")
    print(f"  {'coluna':<28} {'preenchidas':>11} {'atualizadas':>11}")
    for col, (f, u) in stats.items():
        print(f"  {col:<28} {f:>11} {u:>11}")


def _supplement(base: pd.DataFrame, supp: pd.DataFrame, mask: pd.Series) -> None:
    """Preenche, in-place, colunas INFO_COLS de ``base`` nas linhas em ``mask``
    usando valores de ``supp`` indexado por codigo_ies."""
    codes = base.loc[mask, "codigo_ies"]
    for col in INFO_COLS:
        if col in supp.columns:
            base.loc[mask, col] = codes.map(supp[col])


def run() -> None:
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    base = _load_base()

    # Tier 1: o que a base do e-MEC já trouxe → complemento='n'.
    has_base_info = base[INFO_COLS].notna().any(axis=1)
    base["complemento"] = has_base_info.map({True: "n", False: "y"})

    # Tier 2: suplemento do Censo para IES sem info na base.
    missing_after_base = ~has_base_info
    if missing_after_base.any():
        supp_censo = _load_supplement().set_index("codigo_ies")
        _supplement(base, supp_censo, missing_after_base)

    # Tier 3: para IES que continuam sem nada, tenta indicadores_consolidados.
    has_info_after_censo = base[INFO_COLS].notna().any(axis=1)
    missing_after_censo = ~has_info_after_censo
    if missing_after_censo.any():
        supp_indic = _load_indicadores().set_index("codigo_ies")
        _supplement(base, supp_indic, missing_after_censo)
        # As linhas que ganharam dados agora vieram dos indicadores.
        gained_from_indic = base[INFO_COLS].notna().any(axis=1) & missing_after_censo
        base.loc[gained_from_indic, "complemento"] = "i"

    out = base[OUTPUT_COLS].copy()

    before_lengths = out["nome"].dropna().str.len()
    out["nome"] = out["nome"].map(_clean_nome)
    after_lengths = out["nome"].dropna().str.len()
    chars_removed = int(before_lengths.sum() - after_lengths.sum())

    out.to_excel(FINAL_XLSX, index=False)

    has_info = out[INFO_COLS].notna().any(axis=1)
    n_count = (out["complemento"] == "n").sum()
    y_count = (out["complemento"] == "y").sum()
    i_count = (out["complemento"] == "i").sum()
    still_empty = (~has_info).sum()
    print(f"Arquivo gerado: {FINAL_XLSX}")
    print(f"  Total: {len(out)} linhas")
    print(f"  complemento='n' (e-MEC):       {n_count}")
    print(f"  complemento='y' (Censo):       {y_count}")
    print(f"  complemento='i' (Indicadores): {i_count}")
    print(f"  Ainda sem dados: {still_empty}")
    print(f"  Limpeza de nomes: {chars_removed} caracteres removidos.")

    _report_long_names(out)


def run_apply() -> None:
    """Lê ``list_ies_final.xlsx`` do disco e aplica em
    ``indicadores_consolidados.xlsx``. Usado como etapa separada quando
    ``list_ies_final`` já existe."""
    if not FINAL_XLSX.exists():
        print(f"{FINAL_XLSX.name} não encontrado — rode `ies.final.run` antes.")
        return
    apply_to_indicadores(pd.read_excel(FINAL_XLSX))


if __name__ == "__main__":
    run()
    apply_to_indicadores(pd.read_excel(FINAL_XLSX))
