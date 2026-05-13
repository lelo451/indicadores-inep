# -*- coding: utf-8 -*-
"""Funções compartilhadas para consolidar dicionários de variáveis do INEP.

Usado por consolidar_dicionario_enade.py e consolidar_dicionario_idd.py.
"""

from __future__ import annotations

import glob
import os
import re
import unicodedata
import warnings

import pandas as pd

warnings.filterwarnings("ignore")

VARIABLE_PREFIXES = ("CO_", "NU_", "TP_", "NT_", "IN_", "ANO_", "ENEM_", "ID_", "QE_", "DS_")
SINGLE_PAIR_RE = re.compile(r"^\s*(\d+)\s*[=.]\s*(.+?)\s*$")
HEADER_TOKENS = {"NOME", "TIPO", "TAMANHO", "DESCRIÇÃO", "CATEGORIAS", "VARIÁVEL"}

# Reescreve rótulos do ENADE 2008/2011 'ENGENHARIA (GRUPO X) - <CURSO>' para
# o nome moderno e colapsa o agregado 'ENGENHARIA (GRUPO X)' em 'ENGENHARIA'
# antes da deduplicação, evitando linhas duplicadas para o mesmo código.
ENGENHARIA_GRUPO_RE = re.compile(
    r"^\s*ENGENHARIA\s*\(GRUPO\s+[IVX]+\)\s*(?:-\s*(.+?))?\s*$",
    re.IGNORECASE,
)


def drop_engenharia_grupo(value) -> str:
    if pd.isna(value):
        return value
    s = str(value).strip()
    match = ENGENHARIA_GRUPO_RE.match(s)
    if not match:
        return s
    sub = match.group(1)
    return sub.strip() if sub else "ENGENHARIA"


def find_dictionary_files(directory: str, pattern: str) -> list[tuple[int, str]]:
    files = []
    for path in glob.glob(os.path.join(directory, pattern)):
        if "~$" in path:
            continue
        match = re.search(r"(\d{4})", os.path.basename(path))
        if match:
            files.append((int(match.group(1)), path))
    return sorted(files)


def pick_dict_sheet(xl: pd.ExcelFile) -> str | None:
    candidates = [s for s in xl.sheet_names if "DICION" in s.upper()]
    # Prefer the variables sheet over the file-list sheet (DICIONÁRIO_ARQUIVOS).
    for sheet in candidates:
        up = sheet.upper()
        if "VARI" in up:
            return sheet
    for sheet in candidates:
        if "ARQUIVO" not in sheet.upper():
            return sheet
    return candidates[0] if candidates else None


def looks_like_variable_name(token: str) -> bool:
    token = token.strip()
    if not token or " " in token:
        return False
    if token.upper() in HEADER_TOKENS:
        return False
    return any(token.upper().startswith(p) for p in VARIABLE_PREFIXES)


def detect_columns(df: pd.DataFrame) -> tuple[int, int]:
    """Detecta a coluna que contém nomes de variáveis e a de categorias."""
    name_col = None
    for col in (0, 1):
        if col >= df.shape[1]:
            continue
        for value in df.iloc[:, col].dropna():
            token = str(value).strip().upper()
            if token.startswith("CO_") and " " not in token:
                name_col = col
                break
        if name_col is not None:
            break
    if name_col is None:
        name_col = 0

    cat_col = None
    for i in range(min(5, len(df))):
        for j, value in enumerate(df.iloc[i].astype(str)):
            if "CATEGORIA" in value.upper():
                cat_col = j
                break
        if cat_col is not None:
            break
    if cat_col is None:
        cat_col = min(df.shape[1] - 1, 5)
    return name_col, cat_col


def parse_dictionary(
    path: str, targets: set[str], aliases: dict[str, str]
) -> dict[str, list[str]]:
    xl = pd.ExcelFile(path)
    sheet = pick_dict_sheet(xl)
    if sheet is None:
        return {}
    df = pd.read_excel(path, sheet_name=sheet, header=None)
    name_col, cat_col = detect_columns(df)

    out: dict[str, list[str]] = {}
    current: str | None = None

    for _, row in df.iterrows():
        name_cell = row.iloc[name_col] if name_col < len(row) else None
        if isinstance(name_cell, str):
            token = name_cell.strip()
            if token:
                canon = aliases.get(token.upper(), token.upper())
                if canon in targets:
                    current = canon
                    out.setdefault(current, [])
                elif looks_like_variable_name(token):
                    current = None

        if current is None or cat_col >= len(row):
            continue
        cell = row.iloc[cat_col]
        if pd.notna(cell):
            text = str(cell).strip()
            if text:
                out[current].append(text)
    return out


def extract_pairs(text: str) -> list[tuple[str, str]]:
    """Retorna todos os pares (código, rótulo) presentes em uma célula.

    Trata três formatos comuns nos dicionários do INEP:
    - Uma célula por categoria: "1 = Pública Federal"
    - Várias categorias separadas por '\\n' numa única célula
    - Várias categorias na mesma linha, separadas por 2+ espaços (estilo IDD 2014)
    """
    pairs = []
    for line in str(text).split("\n"):
        for segment in re.split(r"\s{2,}", line.strip()):
            seg = segment.strip()
            if not seg:
                continue
            match = SINGLE_PAIR_RE.match(seg)
            if match:
                pairs.append((match.group(1), match.group(2).strip()))
    return pairs


def load_municipios(csv_path: str) -> list[tuple[str, str]]:
    df = pd.read_csv(csv_path, dtype=str)
    code_col = next(c for c in df.columns if "CÓDIGO" in c.upper() and "IBGE" in c.upper())
    name_col = next(
        c for c in df.columns
        if "MUNIC" in c.upper() and "IBGE" in c.upper() and "CÓDIGO" not in c.upper()
    )
    pairs = (
        df[[code_col, name_col]]
        .dropna()
        .drop_duplicates(subset=code_col)
        .itertuples(index=False, name=None)
    )
    return [(code.strip(), name.strip()) for code, name in pairs]


def dedupe_distinct_names(long_df: pd.DataFrame) -> pd.DataFrame:
    """Mantém nomes distintos por variável.

    - Compara descrições por casefold + remoção de acentos.
    - Quando o mesmo nome aparece em vários anos, mantém o ano mais recente e
      usa a grafia daquele ano como canônica.
    - Quando uma descrição tem mais de um código, mantém todos os códigos.
    """

    def normalize(text: str) -> str:
        text = unicodedata.normalize("NFKD", str(text))
        text = "".join(c for c in text if not unicodedata.combining(c))
        return re.sub(r"\s+", " ", text).casefold().strip()

    df = long_df.copy()
    df["descricao"] = df["descricao"].map(drop_engenharia_grupo)
    df["_norm"] = df["descricao"].map(normalize)

    latest = (
        df.sort_values("ano")
        .groupby(["variavel", "_norm"], as_index=False)
        .tail(1)[["variavel", "_norm", "descricao", "ano"]]
        .rename(columns={"descricao": "descricao_canonica", "ano": "ano_mais_recente"})
    )

    codes = df[["variavel", "_norm", "codigo"]].drop_duplicates()

    merged = codes.merge(latest, on=["variavel", "_norm"])
    return (
        merged[["variavel", "codigo", "descricao_canonica", "ano_mais_recente"]]
        .rename(columns={"descricao_canonica": "descricao", "ano_mais_recente": "ano"})
        .sort_values(["variavel", "codigo"])
        .reset_index(drop=True)
    )


def build_long_table(
    files: list[tuple[int, str]],
    variables: list[str],
    aliases: dict[str, str],
    municipios: list[tuple[str, str]],
    municipio_variable: str,
    modalidade_fallback: list[str],
) -> pd.DataFrame:
    targets = set(variables)
    aliases_upper = {k.upper(): v.upper() for k, v in aliases.items()}

    rows = []
    for year, path in files:
        parsed = parse_dictionary(path, targets, aliases_upper)

        if "CO_MODALIDADE" in targets and not parsed.get("CO_MODALIDADE"):
            parsed["CO_MODALIDADE"] = list(modalidade_fallback)

        for variable in variables:
            if variable == municipio_variable and municipios:
                for code, label in municipios:
                    rows.append(
                        {"ano": year, "variavel": variable, "codigo": code, "descricao": label}
                    )
                continue
            for category in parsed.get(variable, []):
                for code, label in extract_pairs(category):
                    rows.append(
                        {"ano": year, "variavel": variable, "codigo": code, "descricao": label}
                    )
    return pd.DataFrame(rows)


def consolidate(
    *,
    dict_dir: str,
    dict_pattern: str,
    variables: list[str],
    output_path: str,
    aliases: dict[str, str] | None = None,
    municipios_csv: str | None = None,
    municipio_variable: str = "CO_MUNIC_CURSO",
    modalidade_fallback: tuple[str, ...] = ("1 = Presencial",),
) -> pd.DataFrame:
    aliases = aliases or {}
    municipios = (
        load_municipios(municipios_csv)
        if municipios_csv and municipio_variable in variables
        else []
    )

    files = find_dictionary_files(dict_dir, dict_pattern)
    long_df = build_long_table(
        files,
        variables=variables,
        aliases=aliases,
        municipios=municipios,
        municipio_variable=municipio_variable,
        modalidade_fallback=list(modalidade_fallback),
    )
    distinct_df = dedupe_distinct_names(long_df)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        distinct_df.to_excel(writer, sheet_name="TODAS", index=False)
        for variable in variables:
            subset = (
                distinct_df[distinct_df["variavel"] == variable]
                .drop(columns="variavel")
                .reset_index(drop=True)
            )
            subset.to_excel(writer, sheet_name=variable, index=False)

    print(f"Saved: {output_path}")
    print(f"Distinct rows: {len(distinct_df)}")
    print("\nRows per variable:")
    print(distinct_df.groupby("variavel").size().to_string())
    return distinct_df
