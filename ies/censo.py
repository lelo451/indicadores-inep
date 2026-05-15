# -*- coding: utf-8 -*-
"""Consolida as listas anuais de IES do Censo da Educação Superior em um único
arquivo (``outputs/lista_ies_consolidada.xlsx``).

Origem dos dados (pasta "Lista INEP Censo Superior/"):
- 2020: já traz as 6 colunas finais (Código, Nome, Sigla, UF, Categoria, Org.).
- 2021/2022/2023: vêm apenas com CO_IES, NO_IES, SG_IES, com uma linha de
  banner no topo (header real está na segunda linha).

Filtragem:
- Apenas IES cujo código esteja em list_ies.csv (coluna codigo_ies) são
  incluídas no resultado; as demais são descartadas antes da consolidação.

Saída:
- Aba "IES": lista deduplicada com as 6 colunas; 2020 é a base, e as IES que
  só aparecem em 2021+ são acrescentadas (UF/Categoria/Organização ficam
  vazias para essas, pois não existem em 2020).
- Aba "Variações": para cada IES cujo nome ou sigla difere entre os anos,
  registra os valores alternativos (separados por ';') em relação à base
  de 2020.
"""

from __future__ import annotations

from collections import OrderedDict

import pandas as pd

from ._common import (
    CONSOLIDADA_XLSX,
    OUTPUTS_DIR,
    SRC_DIR,
    WHITELIST_CSV,
    clean_text,
    normalize_code,
)


YEARS_FULL = [2020]
YEARS_SHORT = [2021, 2022, 2023]

FINAL_COLS = [
    "Codigo da IES",
    "Nome da IES",
    "Sigla da IES",
    "UF",
    "Categoria Administrativa",
    "Organização Acadêmica",
]

SHORT_RENAME = {
    "CO_IES": "Codigo da IES",
    "NO_IES": "Nome da IES",
    "SG_IES": "Sigla da IES",
}


def _read_full(year: int) -> pd.DataFrame:
    path = SRC_DIR / f"Lista_de_IES_{year}.xlsx"
    df = pd.read_excel(path)
    return df[FINAL_COLS].copy()


def _read_short(year: int) -> pd.DataFrame:
    path = SRC_DIR / f"Lista_de_IES_{year}.xlsx"
    # 2023 usa "Sheet0"; 2021/2022 usam "Exportar Planilha".
    sheet = "Sheet0" if year == 2023 else "Exportar Planilha"
    df = pd.read_excel(path, sheet_name=sheet, header=1)
    df = df.rename(columns=SHORT_RENAME)
    return df[["Codigo da IES", "Nome da IES", "Sigla da IES"]].copy()


def _load_whitelist() -> set[int]:
    df = pd.read_csv(WHITELIST_CSV)
    codes = pd.to_numeric(df["codigo_ies"], errors="coerce").dropna().astype(int)
    return set(codes.tolist())


def build_master(per_year: dict[int, pd.DataFrame]) -> pd.DataFrame:
    """Lista deduplicada: 2020 é a base; IES novas dos anos seguintes são
    acrescentadas (sem UF/Categoria/Org, pois não existem em 2020)."""
    base = per_year[2020].copy()
    seen = set(base["Codigo da IES"].dropna().astype(int).tolist())

    extras = []
    for year in YEARS_SHORT:
        df = per_year[year]
        new = df[~df["Codigo da IES"].astype("Int64").isin(seen)]
        if not new.empty:
            extras.append(new.assign(
                **{
                    "UF": pd.NA,
                    "Categoria Administrativa": pd.NA,
                    "Organização Acadêmica": pd.NA,
                }
            )[FINAL_COLS])
            seen.update(new["Codigo da IES"].dropna().astype(int).tolist())

    if extras:
        master = pd.concat([base, *extras], ignore_index=True)
    else:
        master = base

    return master.sort_values("Codigo da IES").reset_index(drop=True)


def build_variations(per_year: dict[int, pd.DataFrame]) -> pd.DataFrame:
    """Para cada código de IES, lista nomes/siglas que aparecem em 2021-2023
    e divergem da base de 2020 (ou da primeira ocorrência, se a IES não
    estiver em 2020). A linha só aparece se houver alguma divergência."""
    by_code: dict[int, "OrderedDict[int, tuple[str, str]]"] = {}
    for year in [2020, *YEARS_SHORT]:
        df = per_year[year]
        for _, row in df.iterrows():
            code = row["Codigo da IES"]
            if pd.isna(code):
                continue
            code = int(code)
            nome = row["Nome da IES"]
            sigla = row["Sigla da IES"]
            by_code.setdefault(code, OrderedDict())[year] = (
                nome if not pd.isna(nome) else "",
                sigla if not pd.isna(sigla) else "",
            )

    rows = []
    for code, entries in by_code.items():
        base_year = 2020 if 2020 in entries else next(iter(entries))
        base_nome, base_sigla = entries[base_year]

        alt_nomes: list[str] = []
        alt_siglas: list[str] = []
        for year, (nome, sigla) in entries.items():
            if year == base_year:
                continue
            if nome and nome != base_nome and nome not in alt_nomes:
                alt_nomes.append(nome)
            if sigla and sigla != base_sigla and sigla not in alt_siglas:
                alt_siglas.append(sigla)

        if alt_nomes or alt_siglas:
            rows.append({
                "cod ies": code,
                "nome ies": ";".join(alt_nomes),
                "sig ies": ";".join(alt_siglas),
            })

    out = pd.DataFrame(rows, columns=["cod ies", "nome ies", "sig ies"])
    return out.sort_values("cod ies").reset_index(drop=True)


def _normalize_year_frame(df: pd.DataFrame, whitelist: set[int]) -> pd.DataFrame:
    df["Codigo da IES"] = normalize_code(df["Codigo da IES"])
    df["Nome da IES"] = clean_text(df["Nome da IES"])
    df["Sigla da IES"] = clean_text(df["Sigla da IES"])
    return df[df["Codigo da IES"].isin(whitelist)].reset_index(drop=True)


def run() -> None:
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    whitelist = _load_whitelist()

    per_year: dict[int, pd.DataFrame] = {}
    for year in YEARS_FULL:
        per_year[year] = _normalize_year_frame(_read_full(year), whitelist)
    for year in YEARS_SHORT:
        per_year[year] = _normalize_year_frame(_read_short(year), whitelist)

    master = build_master(per_year)
    variations = build_variations(per_year)

    with pd.ExcelWriter(CONSOLIDADA_XLSX, engine="openpyxl") as writer:
        master.to_excel(writer, sheet_name="IES", index=False)
        variations.to_excel(writer, sheet_name="Variações", index=False)

    print(f"Arquivo gerado: {CONSOLIDADA_XLSX}")
    print(f"  Aba 'IES': {len(master)} linhas")
    print(f"  Aba 'Variações': {len(variations)} linhas")


if __name__ == "__main__":
    run()
