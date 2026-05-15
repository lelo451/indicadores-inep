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
    MUNICIPIOS_CSV,
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

    Detecta o schema do arquivo:
    - Se houver aba ``IES`` (saída de ``apply_to_indicadores``), lê dela
      diretamente — a dimensão já está normalizada.
    - Caso contrário (saída crua de ``consolidar_indicadores``), lê a aba
      única, agrega com ``first()`` por código de IES, e aplica
      ``ORG_ACAD_LEGACY_CODES`` para sanar códigos numéricos antigos
      (ENADE 2004-2008) que escapam ao decodificador do consolidador.
    """
    xl = pd.ExcelFile(INDICADORES_XLSX)
    if "IES" in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name="IES")
        df = df.rename(columns=INDIC_RENAME)
        df["codigo_ies"] = pd.to_numeric(df["codigo_ies"], errors="coerce").astype("Int64")
        df = df.dropna(subset=["codigo_ies"])
        return df[["codigo_ies", *INFO_COLS]].copy()

    df = pd.read_excel(xl, sheet_name=xl.sheet_names[0])
    df = df.rename(columns=INDIC_RENAME)
    df["codigo_ies"] = pd.to_numeric(df["codigo_ies"], errors="coerce").astype("Int64")
    df = df.dropna(subset=["codigo_ies"])
    df["organizacao"] = df["organizacao"].map(
        lambda v: ORG_ACAD_LEGACY_CODES.get(str(v).strip(), v) if pd.notna(v) else v
    )
    return df.groupby("codigo_ies", as_index=False)[INFO_COLS].first()


_IES_SHEET_RENAME = {
    "codigo_ies": "Código da IES",
    "sigla": "Sigla da IES",
    "nome": "Nome da IES",
    "organizacao": "Organização Acadêmica",
    "categoria": "Categoria Administrativa",
}

_IES_SHEET_COLS = list(_IES_SHEET_RENAME.values())

_IES_DROP_FROM_DADOS = [
    "Sigla da IES",
    "Organização Acadêmica",
    "Categoria Administrativa",
]

_MUNICIPIO_SHEET_COLS = [
    "Código do Município",
    "Município do Curso",
    "Sigla da UF",
]

_MUNICIPIO_DROP_FROM_DADOS = ["Sigla da UF"]


def _build_municipio_dimension(dados: pd.DataFrame) -> pd.DataFrame:
    """Dimensão Município: uma linha por Código do Município presente em
    Dados, com Nome e UF puxados do catálogo IBGE
    (``Microdados/municipios.csv``)."""
    if "Código do Município" not in dados.columns:
        return pd.DataFrame(columns=_MUNICIPIO_SHEET_COLS)

    mun = pd.read_csv(MUNICIPIOS_CSV, dtype=str)
    code_col = next(c for c in mun.columns
                    if "CÓDIGO" in c.upper() and "IBGE" in c.upper())
    name_col = next(c for c in mun.columns
                    if "MUNIC" in c.upper() and "IBGE" in c.upper()
                    and "CÓDIGO" not in c.upper())
    code_to_name = dict(zip(mun[code_col].str.strip(),
                            mun[name_col].fillna("").str.strip()))
    code_to_uf = dict(zip(mun[code_col].str.strip(),
                          mun["UF"].fillna("").str.strip()))

    def _norm_code(value):
        if pd.isna(value):
            return None
        s = str(value).strip()
        return s[:-2] if s.endswith(".0") else s

    used = (
        dados["Código do Município"]
        .dropna()
        .map(_norm_code)
        .dropna()
        .drop_duplicates()
        .sort_values()
        .reset_index(drop=True)
    )
    out = pd.DataFrame({"Código do Município": used})
    out["Município do Curso"] = out["Código do Município"].map(code_to_name)
    out["Sigla da UF"] = out["Código do Município"].map(code_to_uf)
    out = out[
        out["Município do Curso"].notna() & (out["Município do Curso"] != "")
    ].reset_index(drop=True)
    out["Código do Município"] = pd.to_numeric(
        out["Código do Município"], errors="coerce"
    ).astype("Int64")
    return out[_MUNICIPIO_SHEET_COLS]


def _build_ies_dimension(final_df: pd.DataFrame) -> pd.DataFrame:
    """Constrói a dimensão IES a partir de ``list_ies_final``: aplica as
    normalizações canônicas, descarta ``complemento`` e renomeia colunas
    para os rótulos públicos. Sigla=='0' é convertida em NaN (sentinel)."""
    ies = final_df.copy()
    sigla = ies["sigla"].astype("string").str.strip()
    ies.loc[sigla.isin(["0", "0.0"]), "sigla"] = pd.NA
    ies["sigla"] = ies["sigla"].map(strip_double_quotes)
    ies["nome"] = ies["nome"].map(title_case_with_connectors)
    ies["organizacao"] = ies["organizacao"].map(
        lambda v: map_canonical(v, ORG_ACAD_CANONICAL)
    )
    ies["categoria"] = (
        ies["categoria"]
        .map(lambda v: map_canonical(v, CATEG_ADMIN_CANONICAL))
        .map(strip_double_quotes)
    )
    ies["codigo_ies"] = pd.to_numeric(ies["codigo_ies"], errors="coerce").astype("Int64")
    ies = ies.dropna(subset=["codigo_ies"])
    ies = ies.rename(columns=_IES_SHEET_RENAME)
    return (
        ies[_IES_SHEET_COLS]
        .sort_values("Código da IES")
        .reset_index(drop=True)
    )


def apply_to_indicadores(final_df: pd.DataFrame) -> None:
    """Reorganiza ``indicadores_consolidados.xlsx`` em três abas:

    - **Dados**: granularidade de curso. Mantém Código da IES + Nome da IES
      (refrescado da aba IES) e Código do Município + Município do Curso
      (refrescado da aba Municípios). Sigla/Organização/Categoria (IES) e
      Sigla da UF (município) ficam apenas nas respectivas dimensões.
    - **IES**: dimensão por Código da IES (sigla/nome/organização/categoria).
    - **Municípios**: dimensão por Código do Município (nome + UF), do
      catálogo IBGE, restrita aos códigos presentes em Dados.

    Idempotente.
    """
    if not INDICADORES_XLSX.exists():
        print(f"  {INDICADORES_XLSX.name} não encontrado — pulando.")
        return

    ies = _build_ies_dimension(final_df)

    xl = pd.ExcelFile(INDICADORES_XLSX)
    dados_sheet = "Dados" if "Dados" in xl.sheet_names else xl.sheet_names[0]
    dados = pd.read_excel(xl, sheet_name=dados_sheet)

    municipios = _build_municipio_dimension(dados)

    nome_lookup = dict(zip(ies["Código da IES"], ies["Nome da IES"]))
    dados["Código da IES"] = pd.to_numeric(
        dados["Código da IES"], errors="coerce"
    ).astype("Int64")
    mapped_nome = dados["Código da IES"].map(nome_lookup)
    ies_nome_hits = mapped_nome.notna()
    dados.loc[ies_nome_hits, "Nome da IES"] = mapped_nome[ies_nome_hits]

    mun_nome_hits = pd.Series(False, index=dados.index)
    if "Código do Município" in dados.columns and not municipios.empty:
        dados["Código do Município"] = pd.to_numeric(
            dados["Código do Município"], errors="coerce"
        ).astype("Int64")
        mun_lookup = dict(zip(
            municipios["Código do Município"],
            municipios["Município do Curso"],
        ))
        mapped_mun = dados["Código do Município"].map(mun_lookup)
        mun_nome_hits = mapped_mun.notna()
        dados.loc[mun_nome_hits, "Município do Curso"] = mapped_mun[mun_nome_hits]

    drop_cols = _IES_DROP_FROM_DADOS + _MUNICIPIO_DROP_FROM_DADOS
    dados = dados.drop(columns=[c for c in drop_cols if c in dados.columns])

    with pd.ExcelWriter(INDICADORES_XLSX, engine="openpyxl") as writer:
        dados.to_excel(writer, sheet_name="Dados", index=False)
        ies.to_excel(writer, sheet_name="IES", index=False)
        municipios.to_excel(writer, sheet_name="Municípios", index=False)

    print(f"Reorganizado em {INDICADORES_XLSX.name}:")
    print(f"  Aba 'Dados':      {len(dados):>6} linhas × {len(dados.columns):>2} colunas")
    print(f"  Aba 'IES':        {len(ies):>6} linhas × {len(ies.columns):>2} colunas")
    print(f"  Aba 'Municípios': {len(municipios):>6} linhas × {len(municipios.columns):>2} colunas")
    print(
        f"  Nome da IES refrescado em {int(ies_nome_hits.sum())} linhas; "
        f"Município do Curso em {int(mun_nome_hits.sum())} linhas"
    )


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
