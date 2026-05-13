# -*- coding: utf-8 -*-
"""Consolida os arquivos de indicadores (Enade, CPC, IDD, IGC, enade_cpc) em
uma única planilha com colunas canônicas.

Estratégia:
- Cada arquivo em data/ contribui para um subconjunto de colunas conforme seu
  tipo (Enade contribui Enade Faixa/Contínuo; CPC contribui CPC ...; IDD idem;
  IGC contribui IGC Faixa/Contínuo).
- IGC pré-2017 vem dividido em múltiplas planilhas por organização acadêmica
  (Universidades, Centros, Faculdades); injetamos a Organização a partir do
  nome da planilha.
- Planilhas auxiliares (Atualizações, Programas Capes, matrículas, etc.) são
  filtradas.
- Valores codificados (co_grupo=5, co_categad=1 etc.) são decodificados pelos
  dicionários consolidados (dicionario_enade_consolidado.xlsx / _idd_).
- O resultado é a granularidade de curso (Ano + Código do Curso). IGC é
  agregado na granularidade de IES (Ano + Código da IES) e mesclado por
  LEFT JOIN.
"""

from __future__ import annotations

import glob
import gzip
import json
import os
import re
import unicodedata
import urllib.error
import urllib.request
import warnings

import pandas as pd

warnings.filterwarnings("ignore")

DATA_ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(DATA_ROOT, "data")
DICT_ENADE = os.path.join(DATA_ROOT, "dicionario_enade_consolidado.xlsx")
DICT_IDD = os.path.join(DATA_ROOT, "dicionario_idd_consolidado.xlsx")
OUTPUT = os.path.join(DATA_ROOT, "indicadores_consolidados.xlsx")

TARGETS = [
    "Ano", "Código da Área", "Área de Avaliação", "Grau Acadêmico",
    "Código da IES", "Nome da IES", "Sigla da IES",
    "Organização Acadêmica", "Categoria Administrativa",
    "Código do Curso", "Modalidade de Ensino",
    "Código do Município", "Município do Curso", "Sigla da UF",
    "Conceito Enade (Contínuo)", "Conceito Enade (Faixa)",
    "IGC (Contínuo)", "IGC (Faixa)",
    "CPC (Contínuo)", "CPC (Faixa)",
    "IDD (Contínuo)", "IDD (Faixa)",
]

COLUMN_ALIASES = {
    "Ano": ["Ano", "Ano Enade", "Ano CPC", "Ano IGC", "Ano IDD", "nu_ano",
            "ultimo ano avaliado no enade", "Edição"],
    "Código da Área": ["Código da Área", "Cód Area", "Cód.Área", "Cod. Area",
                       "Cod Area", "co_grupo", "co_subarea", "Cod. da Área",
                       "Código Área"],
    "Área de Avaliação": ["Área de Avaliação", "Área", "Área de Enquadramento",
                          "Área Enquadramento", "Nome da Área", "Sub Area",
                          "Descrição da Área"],
    "Grau Acadêmico": ["Grau Acadêmico", "Grau"],
    "Código da IES": ["Código da IES", "Cód. IES", "Cód.IES", "Código_da_IES",
                      "co_ies", "Cod. IES", "Código IES"],
    "Nome da IES": ["Nome da IES", "IES", "no_ies"],
    "Sigla da IES": ["Sigla da IES", "Sigla IES", "Sigla", "sg_ies"],
    "Organização Acadêmica": ["Organização Acadêmica", "Org. Acadêmica",
                              "Organização_Acadêmica", "Organização",
                              "co_orgacad", "co_orgac", "cd_orgac", "no_orgacad"],
    "Categoria Administrativa": ["Categoria Administrativa", "Categ. Administrativa",
                                 "Categoria_Administrativa", "Dependência Administrativa",
                                 "Dep. Administrativa", "co_categad", "co_catad",
                                 "cd_catad", "cd_dep", "Rede", "no_categad"],
    "Código do Curso": ["Código do Curso", "co_curso", "Cod. Curso"],
    "Modalidade de Ensino": ["Modalidade de Ensino", "Modalidade", "co_modalidade"],
    "Código do Município": ["Código do Município", "Cod. Município", "co_munic",
                            "co_munic_curso", "codmunic_inep", "Cod Município",
                            "Código Município", "Código do Município do Curso"],
    "Município do Curso": ["Município do Curso", "Município (funcionamento do curso)",
                           "Nome do Município", "Município", "Município_Sede",
                           "Município (Sede)"],
    "Sigla da UF": ["Sigla da UF", "UF", "UF (Sede)", "UF do Curso", "UF da IES",
                    "UF_Sede", "co_uf_curso", "sg_uf", "Sigla UF", "Código UF"],
    "Conceito Enade (Contínuo)": ["Conceito Enade (Contínuo)", "Enade Contínuo",
                                  "Conceito Enade Contínuo", "Nota Contínua Enade",
                                  "Nota Contínua do Enade", "Nota Enade Concluintes",
                                  "Nota Enade Concluintes = Conceito Enade Contínuo"],
    "Conceito Enade (Faixa)": ["Conceito Enade (Faixa)", "Enade Faixa",
                               "Conceito_Enade", "Conceito Enade Faixa",
                               "Conceito Enade"],
    "IGC (Contínuo)": ["IGC (Contínuo)", "IGC Contínuo", "IGC - Contínuo",
                       "IGC", "igc"],
    "IGC (Faixa)": ["IGC (Faixa)", "IGC Faixa", "IGC - Faixas", "IGC - Faixa",
                    "fx_igc"],
    "CPC (Contínuo)": ["CPC (Contínuo)", "CPC Contínuo", "CPC - Contínuo",
                       "Nota Contínua CPC"],
    "CPC (Faixa)": ["CPC (Faixa)", "CPC Faixa", "CPC - Faixa"],
    "IDD (Contínuo)": ["IDD (Contínuo)", "IDD Contínuo", "Nota Contínua IDD",
                       "Nota IDD", "Nota Padronizada - IDD"],
    "IDD (Faixa)": ["IDD (Faixa)", "IDD Faixa", "Conceito_IDD"],
}

# Sheets we should never read in any file
SHEET_SKIP_RE = re.compile(
    r"atualiz|programa|capes|versão|versões|versao|leia|dicion|"
    r"matr[ií]cula|"
    r"^plan(ilha)?\s*[23]$|^sheet\s*[23]$",
    re.IGNORECASE,
)

# Additional skips that apply only inside IGC files (auxiliary slices)
IGC_AUX_SKIP_RE = re.compile(
    r"cursos com cpc|unidades com cpc|"
    r"^cpc[_\s]\d{4}[_\s]\d{4}|^cpc[_\s]?\d{4}\b",
    re.IGNORECASE,
)


def norm(s) -> str:
    s = re.sub(r"[*]+", "", str(s))
    s = s.replace("_", " ")
    s = re.sub(r"\bcod\.\s*", "cod ", s, flags=re.I)
    s = re.sub(r"\bcateg\.\s*", "categoria ", s, flags=re.I)
    s = re.sub(r"\borg\.\s*", "organizacao ", s, flags=re.I)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().casefold()


ALIAS_MAP = {norm(a): t for t, aliases in COLUMN_ALIASES.items() for a in aliases}

NUMERIC_TARGETS = {
    "Conceito Enade (Contínuo)", "Conceito Enade (Faixa)",
    "IGC (Contínuo)", "IGC (Faixa)",
    "CPC (Contínuo)", "CPC (Faixa)",
    "IDD (Contínuo)", "IDD (Faixa)",
}


def file_kind(name: str) -> str | None:
    base = os.path.basename(name)
    if base.startswith("IGC_"): return "igc"
    if base.startswith("idd_"): return "idd"
    if base.startswith("enade_cpc_"): return "enade_cpc"
    if base.startswith("enade_"): return "enade"
    if base.startswith("cpc_"): return "cpc"
    return None


def year_from_filename(name: str) -> str | None:
    m = re.search(r"(20\d{2}|19\d{2})", os.path.basename(name))
    return m.group(1) if m else None


def igc_org_from_sheet(sheet: str) -> str | None:
    s = sheet.lower()
    if any(x in s for x in ("matr", "capes", "programa", "cpc")):
        return None
    if "centro" in s or "cefet" in s:
        return "Centro Universitário"
    if "universidade" in s or "ifet" in s:
        return "Universidade"
    if "faculdade" in s or "outros" in s:
        return "Faculdade"
    return None


def pick_data_sheets(path: str, kind: str) -> list[str]:
    xl = pd.ExcelFile(path)
    candidates = []
    for s in xl.sheet_names:
        if SHEET_SKIP_RE.search(s):
            continue
        if kind == "igc" and IGC_AUX_SKIP_RE.search(s):
            continue
        try:
            raw = pd.read_excel(path, sheet_name=s, header=None, nrows=200, dtype=str)
        except Exception:
            continue
        if raw.dropna(how="all").shape[0] < 5:
            continue
        candidates.append(s)

    if kind == "igc":
        org_sheets = [s for s in candidates if igc_org_from_sheet(s) is not None]
        if org_sheets:
            return org_sheets

    kw = {"enade": "ENADE", "cpc": "CPC", "igc": "IGC", "idd": "IDD"}.get(kind)
    if kw:
        by_kw = [s for s in candidates if kw in s.upper()]
        if by_kw:
            return [by_kw[0]]

    not_generic = [s for s in candidates
                   if not re.match(r"(plan(ilha)?\d+|sheet\d+)$", s, re.I)]
    if not_generic:
        return [not_generic[0]]
    return candidates[:1]


def detect_header_row(raw: pd.DataFrame) -> int:
    keywords = ("ano", "ies", "curso", "codigo", "area", "uf", "municipio",
                "grupo", "categad", "modalidade", "categoria", "organizacao",
                "sigla", "nome", "cod")
    best = (0, -1)
    for i in range(min(10, len(raw))):
        row = raw.iloc[i].dropna()
        if len(row) < 3:
            continue
        signal = sum(1 for v in row.astype(str)
                     if any(k in norm(v) for k in keywords))
        if signal > best[1]:
            best = (i, signal)
    return best[0]


def read_sheet(path: str, sheet: str) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=12, dtype=str)
    hr = detect_header_row(raw)
    df = pd.read_excel(path, sheet_name=sheet, header=hr, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename, used_targets = {}, set()
    for c in df.columns:
        cn = norm(c)
        tgt = ALIAS_MAP.get(cn)
        if tgt and tgt not in used_targets:
            rename[c] = tgt
            used_targets.add(tgt)
    df = df.rename(columns=rename)
    keep = [c for c in df.columns if c in TARGETS]
    return df[keep].copy()


def load_decode_dicts() -> dict[str, dict[str, str]]:
    decoders: dict[str, dict[str, str]] = {}
    mapping = {
        "CO_CATEGAD": "Categoria Administrativa",
        "CO_ORGACAD": "Organização Acadêmica",
        "CO_GRUPO": "Área de Avaliação",
        "CO_MODALIDADE": "Modalidade de Ensino",
        "CO_UF_CURSO": "Sigla da UF",
        "CO_MUNIC_CURSO": "Município do Curso",
    }
    for path in (DICT_ENADE, DICT_IDD):
        if not os.path.exists(path):
            continue
        xl = pd.ExcelFile(path)
        for var, tgt in mapping.items():
            if var not in xl.sheet_names:
                continue
            d = pd.read_excel(path, sheet_name=var, dtype=str)
            for _, row in d.iterrows():
                code = str(row["codigo"]).strip()
                desc = str(row["descricao"]).strip()
                decoders.setdefault(tgt, {}).setdefault(code, desc)
    return decoders


def decode_column(series: pd.Series, decoder: dict[str, str]) -> pd.Series:
    def lookup(v):
        if pd.isna(v):
            return v
        key = str(v).strip()
        return decoder.get(key, v)
    return series.map(lookup)


UF_FULL_TO_ABBREV = {
    "ACRE": "AC", "ALAGOAS": "AL", "AMAPA": "AP", "AMAZONAS": "AM",
    "BAHIA": "BA", "CEARA": "CE", "DISTRITO FEDERAL": "DF",
    "ESPIRITO SANTO": "ES", "GOIAS": "GO", "MARANHAO": "MA",
    "MATO GROSSO": "MT", "MATO GROSSO DO SUL": "MS", "MINAS GERAIS": "MG",
    "PARA": "PA", "PARAIBA": "PB", "PARANA": "PR", "PERNAMBUCO": "PE",
    "PIAUI": "PI", "RIO DE JANEIRO": "RJ", "RIO GRANDE DO NORTE": "RN",
    "RIO GRANDE DO SUL": "RS", "RONDONIA": "RO", "RORAIMA": "RR",
    "SANTA CATARINA": "SC", "SAO PAULO": "SP", "SERGIPE": "SE",
    "TOCANTINS": "TO",
}


def _strip_accents(value: str) -> str:
    value = unicodedata.normalize("NFKD", str(value))
    return "".join(c for c in value if not unicodedata.combining(c))


def clean_uf(value):
    if pd.isna(value):
        return value
    s = str(value).strip()
    m = re.search(r"\(([A-Za-z]{2})\)\s*$", s)
    if m:
        return m.group(1).upper()
    if len(s) == 2 and s.isalpha():
        return s.upper()
    s_upper = _strip_accents(s).upper()
    if s_upper in UF_FULL_TO_ABBREV:
        return UF_FULL_TO_ABBREV[s_upper]
    return s


def norm_municipio(value) -> str | None:
    if pd.isna(value):
        return None
    value = _strip_accents(value)
    value = re.sub(r"[^\w\s]", " ", value, flags=re.UNICODE)
    return re.sub(r"\s+", " ", value).strip().upper()


def to_numeric(series: pd.Series) -> pd.Series:
    cleaned = series.astype(str).str.replace(",", ".", regex=False).str.strip()
    return pd.to_numeric(cleaned, errors="coerce")


def process_file(path: str, decoders: dict) -> pd.DataFrame:
    kind = file_kind(path)
    year = year_from_filename(path)
    pieces = []
    for sheet in pick_data_sheets(path, kind):
        df = read_sheet(path, sheet)
        df = normalize_columns(df)
        if df.empty:
            continue
        if "Ano" not in df.columns and year:
            df["Ano"] = year
        if kind == "igc" and "Organização Acadêmica" not in df.columns:
            org = igc_org_from_sheet(sheet)
            if org:
                df["Organização Acadêmica"] = org
        for tgt, decoder in decoders.items():
            if tgt in df.columns:
                df[tgt] = decode_column(df[tgt], decoder)
        if "Sigla da UF" in df.columns:
            df["Sigla da UF"] = df["Sigla da UF"].map(clean_uf)
        for tgt in NUMERIC_TARGETS:
            if tgt in df.columns:
                df[tgt] = to_numeric(df[tgt])
        pieces.append(df)
    if not pieces:
        return pd.DataFrame()
    return pd.concat(pieces, ignore_index=True, sort=False)


def collect_all() -> tuple[list[pd.DataFrame], list[pd.DataFrame]]:
    decoders = load_decode_dicts()
    courses, igc = [], []
    files = sorted(glob.glob(os.path.join(DATA_DIR, "*.xls*")))
    for path in files:
        kind = file_kind(path)
        if not kind:
            continue
        df = process_file(path, decoders)
        if df.empty:
            continue
        (igc if kind == "igc" else courses).append(df)
        print(f"  {os.path.basename(path):40} -> {len(df):6d} rows, {len(df.columns)} cols")
    return courses, igc


def merge_on(df: pd.DataFrame, keys: list[str]) -> pd.DataFrame:
    keys = [k for k in keys if k in df.columns]
    if not keys:
        return df
    agg = {c: "first" for c in df.columns if c not in keys}
    return df.groupby(keys, dropna=False, as_index=False).agg(agg)


# Spellings INEP used that don't match the canonical IBGE name.
# Each entry maps (UF, alternate name as it appears in source files) -> canonical IBGE name.
MUNICIPIO_NAME_ALIASES = {
    ("RJ", "Campos do Goytacazes"):       "Campos dos Goytacazes",
    ("PE", "Jaboatão do Guararapes"):     "Jaboatão dos Guararapes",
    ("SC", "São Miguel d'Oeste"):         "São Miguel do Oeste",
    ("PE", "Belém de São Francisco"):     "Belém do São Francisco",
    ("MT", "São José do Quatro Marcos"):  "São José dos Quatro Marcos",
    ("AL", "Palmeira do Índios"):         "Palmeira dos Índios",
    ("PR", "São José do Pinhais"):        "São José dos Pinhais",
    ("RS", "Santana do Livramento"):      "Sant'Ana do Livramento",
    ("SP", "Ipauçu"):                     "Ipaussu",  # INEP misspelled the IBGE 'Ipaussu'
}


def load_municipio_lookup() -> dict[tuple[str, str], str]:
    """Return {(UF, NOME_NORMALIZADO): codigo} from Microdados/municipios.csv,
    augmented with the known spelling aliases above."""
    path = os.path.join(DATA_ROOT, "Microdados", "municipios.csv")
    if not os.path.exists(path):
        return {}
    df = pd.read_csv(path, dtype=str)
    code_col = next(c for c in df.columns if "CÓDIGO" in c.upper() and "IBGE" in c.upper())
    name_col = next(c for c in df.columns
                    if "MUNIC" in c.upper() and "IBGE" in c.upper() and "CÓDIGO" not in c.upper())
    out = {}
    for _, row in df.iterrows():
        out[(str(row["UF"]).strip().upper(), norm_municipio(row[name_col]))] = str(row[code_col]).strip()
    # Pull codes for the alias targets and register the alternate keys
    for (uf, alt_name), canonical in MUNICIPIO_NAME_ALIASES.items():
        code = out.get((uf, norm_municipio(canonical)))
        if code is not None:
            out[(uf, norm_municipio(alt_name))] = code
    return out


MUNICIPIOS_CSV = os.path.join(os.path.dirname(__file__), "Microdados", "municipios.csv")
IBGE_API = "https://servicodados.ibge.gov.br/api/v1/localidades/municipios/{cod}"


def _normalize_code(value) -> str | None:
    if pd.isna(value):
        return None
    s = str(value).strip()
    return s[:-2] if s.endswith(".0") else s


def _fetch_ibge_record(code: str) -> dict | None:
    """Return {'nome': ..., 'uf': ...} for an IBGE municipality code, or None."""
    req = urllib.request.Request(
        IBGE_API.format(cod=code),
        headers={"Accept-Encoding": "identity", "User-Agent": "consolidar-indicadores/1.0"},
    )
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            raw = resp.read()
            if raw[:2] == b"\x1f\x8b":  # gzip magic bytes
                raw = gzip.decompress(raw)
            data = json.loads(raw)
        uf = (data.get("microrregiao", {})
                  .get("mesorregiao", {})
                  .get("UF", {})
                  .get("sigla", ""))
        nome = data.get("nome")
        if not nome:
            return None
        return {"nome": nome, "uf": uf}
    except (urllib.error.URLError, ValueError, TimeoutError, UnicodeDecodeError) as exc:
        print(f"  IBGE API failed for {code}: {exc}")
        return None


def fill_municipio_name(df: pd.DataFrame) -> pd.DataFrame:
    """Fill Município do Curso from Código do Município. Local lookup first;
    falls back to the IBGE localidades API. Any names resolved via the API are
    appended to municipios.csv so subsequent runs hit the local file."""
    if not {"Código do Município", "Município do Curso"}.issubset(df.columns):
        return df

    mun = pd.read_csv(MUNICIPIOS_CSV, dtype=str)
    code_col = next(c for c in mun.columns if "CÓDIGO" in c.upper() and "IBGE" in c.upper())
    name_col = next(c for c in mun.columns
                    if "MUNIC" in c.upper() and "IBGE" in c.upper() and "CÓDIGO" not in c.upper())
    mun[code_col] = mun[code_col].str.strip()
    code_to_name = {
        c: n for c, n in zip(mun[code_col], mun[name_col].fillna("").str.strip()) if n
    }

    needs = df["Código do Município"].notna() & df["Município do Curso"].isna()
    codes = df.loc[needs, "Código do Município"].map(_normalize_code)
    df.loc[needs, "Município do Curso"] = codes.map(code_to_name)
    local_filled = int(df.loc[needs, "Município do Curso"].notna().sum())

    still_needs = df["Código do Município"].notna() & df["Município do Curso"].isna()
    api_filled, updated, added = 0, 0, 0
    if still_needs.any():
        unresolved = (
            df.loc[still_needs, "Código do Município"].map(_normalize_code).dropna().unique()
        )
        existing = set(mun[code_col])
        for code in unresolved:
            rec = _fetch_ibge_record(code)
            if not rec:
                continue
            code_to_name[code] = rec["nome"]
            if code in existing:
                row_mask = mun[code_col] == code
                mun.loc[row_mask, name_col] = rec["nome"]
                if "UF" in mun.columns and rec.get("uf"):
                    mun.loc[row_mask, "UF"] = mun.loc[row_mask, "UF"].fillna(rec["uf"])
                updated += 1
            else:
                new_row = {code_col: code, name_col: rec["nome"], "UF": rec["uf"]}
                mun = pd.concat([mun, pd.DataFrame([new_row])], ignore_index=True)
                added += 1
        df.loc[still_needs, "Município do Curso"] = (
            df.loc[still_needs, "Código do Município"].map(_normalize_code).map(code_to_name)
        )
        api_filled = int(df.loc[still_needs, "Município do Curso"].notna().sum())
        if updated or added:
            mun.to_csv(MUNICIPIOS_CSV, index=False)
            print(f"  municipios.csv: {updated} updated, {added} appended")

    print(f"  filled {local_filled} from local + {api_filled} from IBGE API")
    return df


def fill_municipio_code(df: pd.DataFrame) -> pd.DataFrame:
    """Fill Código do Município from Município do Curso + Sigla da UF."""
    needed = {"Código do Município", "Município do Curso", "Sigla da UF"}
    if not needed.issubset(df.columns):
        return df
    lookup = load_municipio_lookup()
    if not lookup:
        return df

    uf_norm = df["Sigla da UF"].map(clean_uf).astype("string").str.upper()
    name_norm = df["Município do Curso"].map(norm_municipio)
    keys = list(zip(uf_norm, name_norm))
    resolved = pd.Series([lookup.get(k) for k in keys], index=df.index, dtype="object")
    mask = df["Código do Município"].isna() & resolved.notna()
    df.loc[mask, "Código do Município"] = resolved[mask]
    print(f"  filled {int(mask.sum())} município codes from name+UF lookup")
    return df


def normalize_ies_metadata(df: pd.DataFrame) -> pd.DataFrame:
    """For each Código da IES, overwrite Nome/Sigla/Categoria/Organização
    with the values from the most recent year that has them non-null.
    Groups all rows of the same IES under a single canonical identity."""
    ies_cols = ["Nome da IES", "Sigla da IES",
                "Categoria Administrativa", "Organização Acadêmica"]
    ies_cols = [c for c in ies_cols if c in df.columns]
    if not ies_cols or "Código da IES" not in df.columns or "Ano" not in df.columns:
        return df

    sorted_df = df.sort_values("Ano", ascending=False)
    changes = {}
    for col in ies_cols:
        latest = (
            sorted_df.dropna(subset=[col])
            .drop_duplicates(subset=["Código da IES"], keep="first")
            .set_index("Código da IES")[col]
        )
        changes[col] = latest
    code_series = df["Código da IES"]
    for col, latest in changes.items():
        mapped = code_series.map(latest)
        df[col] = mapped.where(mapped.notna(), df[col])
    return df


def main() -> None:
    print("Loading data files...")
    course_dfs, igc_dfs = collect_all()

    print("\nMerging course-level rows...")
    courses = pd.concat(course_dfs, ignore_index=True, sort=False) if course_dfs else pd.DataFrame()
    if not courses.empty:
        has_curso = courses.get("Código do Curso", pd.Series([None] * len(courses))).notna()
        with_curso = merge_on(courses[has_curso], ["Ano", "Código do Curso"])
        without_curso = merge_on(
            courses[~has_curso],
            ["Ano", "Código da IES", "Código da Área", "Código do Município"],
        )
        courses = pd.concat([with_curso, without_curso], ignore_index=True, sort=False)
        print(f"  course-level rows after merge: {len(courses)}")

    print("\nMerging IGC rows...")
    igc = pd.concat(igc_dfs, ignore_index=True, sort=False) if igc_dfs else pd.DataFrame()
    if not igc.empty:
        igc = merge_on(igc, ["Ano", "Código da IES"])
        print(f"  IES-level rows after merge: {len(igc)}")

    print("\nJoining IGC onto course rows...")
    if not courses.empty and not igc.empty and {"Ano", "Código da IES"}.issubset(courses.columns):
        igc_join_cols = ["Ano", "Código da IES", "IGC (Contínuo)", "IGC (Faixa)"]
        igc_join_cols = [c for c in igc_join_cols if c in igc.columns]
        ies_meta_cols = [c for c in ("Nome da IES", "Sigla da IES",
                                     "Organização Acadêmica", "Categoria Administrativa",
                                     "Sigla da UF") if c in igc.columns]
        merge_cols = list(dict.fromkeys(igc_join_cols + ies_meta_cols))
        courses = courses.merge(
            igc[merge_cols], on=["Ano", "Código da IES"], how="left", suffixes=("", "__igc")
        )
        for c in courses.columns:
            if c.endswith("__igc"):
                base = c[:-5]
                if base in courses.columns:
                    courses[base] = courses[base].fillna(courses[c])
                courses = courses.drop(columns=c)

        # Append IGC-only rows (IES that never appear in course-level files)
        ies_in_courses = courses[["Ano", "Código da IES"]].drop_duplicates()
        merged = igc.merge(ies_in_courses, on=["Ano", "Código da IES"], how="left", indicator=True)
        igc_only = merged[merged["_merge"] == "left_only"].drop(columns="_merge")
        if not igc_only.empty:
            courses = pd.concat([courses, igc_only], ignore_index=True, sort=False)
            print(f"  IGC-only rows added: {len(igc_only)}")

    for c in TARGETS:
        if c not in courses.columns:
            courses[c] = pd.NA
    courses = courses[TARGETS]

    print("\nDropping footnote/header garbage rows...")
    before = len(courses)
    valid_ano = courses["Ano"].astype("string").str.fullmatch(r"\s*(?:19|20)\d{2}\s*")
    valid_ies = courses["Código da IES"].astype("string").str.fullmatch(r"\s*\d+\s*")
    courses = courses[valid_ano.fillna(False) & valid_ies.fillna(False)].reset_index(drop=True)
    print(f"  dropped {before - len(courses)} rows")

    print("\nFilling Código do Município from name + UF where missing...")
    courses = fill_municipio_code(courses)

    print("\nFilling Município do Curso from Código (local then IBGE API)...")
    courses = fill_municipio_name(courses)

    print("\nNormalizing IES metadata to latest year per Código da IES...")
    courses = normalize_ies_metadata(courses)

    sort_keys = [k for k in ("Ano", "Código da IES", "Código do Curso") if k in courses.columns]
    courses = courses.sort_values(sort_keys, na_position="last").reset_index(drop=True)

    print(f"\nWriting {len(courses)} rows to {OUTPUT}...")
    courses.to_excel(OUTPUT, index=False)

    print("\nCoverage:")
    for c in TARGETS:
        n = courses[c].notna().sum()
        print(f"  {c:32}  {n:7d} / {len(courses)}")


if __name__ == "__main__":
    main()
