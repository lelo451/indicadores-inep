# -*- coding: utf-8 -*-
"""Gera ``outputs/list_ies_final.xlsx`` combinando a base do e-MEC
(``Microdados/ies_siglas.csv``) com a lista consolidada do Censo
(``outputs/lista_ies_consolidada.xlsx``).

Regra:
- Para cada IES da base, se ela tiver dados (sigla/nome/organização/categoria),
  usamos a base e marcamos complemento = 'n'.
- Caso a base não tenha dados (match_type = not_found), tentamos suplementar
  com a planilha consolidada do Censo e marcamos complemento = 'y'.

Limpeza:
- Nomes do e-MEC frequentemente trazem anotações administrativas anexadas
  (ex.: "Em supervisão", "Suspensão contrato FIES", "Sub Judice"). O script
  trunca o nome a partir do primeiro desses marcadores. Ao final, imprime
  os nomes ainda mais longos que a média + 2·desvio padrão, para inspeção.
"""

from __future__ import annotations

import re

import pandas as pd

from ._common import BASE_CSV, CONSOLIDADA_XLSX, FINAL_XLSX, OUTPUTS_DIR


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


def run() -> None:
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    base = _load_base()
    supp = _load_supplement().set_index("codigo_ies")

    # 'n' quando a base tem ao menos um campo de informação; 'y' caso contrário.
    has_base_info = base[INFO_COLS].notna().any(axis=1)
    base["complemento"] = has_base_info.map({True: "n", False: "y"})

    # Suplementa apenas as linhas sem info na base, célula a célula.
    missing_mask = ~has_base_info
    if missing_mask.any():
        codes = base.loc[missing_mask, "codigo_ies"]
        for col in INFO_COLS:
            values = codes.map(supp[col]) if col in supp.columns else None
            base.loc[missing_mask, col] = values

    out = base[OUTPUT_COLS].copy()

    before_lengths = out["nome"].dropna().str.len()
    out["nome"] = out["nome"].map(_clean_nome)
    after_lengths = out["nome"].dropna().str.len()
    chars_removed = int(before_lengths.sum() - after_lengths.sum())

    out.to_excel(FINAL_XLSX, index=False)

    suplementadas = (out["complemento"] == "y").sum()
    suplementadas_com_dados = (
        (out["complemento"] == "y") & out[INFO_COLS].notna().any(axis=1)
    ).sum()
    print(f"Arquivo gerado: {FINAL_XLSX}")
    print(f"  Total: {len(out)} linhas")
    print(f"  complemento='n' (da base): {(out['complemento'] == 'n').sum()}")
    print(f"  complemento='y' (suplementadas): {suplementadas}")
    print(f"    destas, com dados após suplemento: {suplementadas_com_dados}")
    print(f"    destas, ainda sem dados: {suplementadas - suplementadas_com_dados}")
    print(f"  Limpeza de nomes: {chars_removed} caracteres removidos.")

    _report_long_names(out)


if __name__ == "__main__":
    run()
