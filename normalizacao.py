# -*- coding: utf-8 -*-
"""Tabelas canônicas e utilitários de normalização compartilhados pelos
consolidadores de indicadores e de dicionários.

Originalmente vivia em ``dicionario_comum.py``, mas só metade das funções era
realmente de domínio "dicionário"; a outra metade (tabelas canônicas + helpers
de texto) é usada também pelo ``consolidar_indicadores.py``. Esta separação
explicita a direção das dependências:

    normalizacao  ←  dicionario  ←  dicionario.{enade,idd}
    normalizacao  ←  consolidar_indicadores
"""

from __future__ import annotations

import re
import unicodedata

import pandas as pd


PORTUGUESE_CONNECTORS = {
    "a", "à", "as", "às", "ao", "aos", "o", "os",
    "da", "das", "de", "do", "dos",
    "e", "em", "na", "nas", "no", "nos",
    "com", "para", "por", "sem", "sob", "sobre",
    "ou", "entre",
}


ORG_ACAD_CANONICAL = {
    "universidade": "Universidade",
    "universidades": "Universidade",
    "centro universitario": "Centro Universitário",
    "centros universitarios": "Centro Universitário",
    "faculdade": "Faculdade",
    "faculdades": "Faculdade",
    "faculdades integradas": "Faculdade",
    "faculdade de tecnologia": "Faculdade de Tecnologia",
    "centro federal de educacao tecnologica": "Centro Federal de Educação Tecnológica",
    "cefet": "Centro Federal de Educação Tecnológica",
    "centro de educacao tecnologica": "Centro de Educação Tecnológica",
    "instituto federal de educacao ciencia e tecnologia": "Instituto Federal de Educação, Ciência e Tecnologia",
    "ifet": "Instituto Federal de Educação, Ciência e Tecnologia",
    "if": "Instituto Federal de Educação, Ciência e Tecnologia",
    "instituto superior ou escola superior": "Instituto Superior ou Escola Superior",
    "instituto superior": "Instituto Superior ou Escola Superior",
    "escola superior": "Instituto Superior ou Escola Superior",
}


CATEG_ADMIN_CANONICAL = {
    "publica federal": "Pública Federal",
    "publica estadual": "Pública Estadual",
    "publica municipal": "Pública Municipal",
    "federal": "Pública Federal",
    "estadual": "Pública Estadual",
    "municipal": "Pública Municipal",
    "publica": "Pública",
    "privada": "Privada",
    "privada com fins lucrativos": "Privada com fins lucrativos",
    "privada sem fins lucrativos": "Privada sem fins lucrativos",
    "especial": "Especial",
    "comunitaria confessional": "Comunitária/Confessional",
    "pessoa juridica de direito publico federal": "Pública Federal",
    "pessoa juridica de direito publico estadual": "Pública Estadual",
    "pessoa juridica de direito publico municipal": "Pública Municipal",
    "pessoa juridica de direito privado com fins lucrativos": "Privada com fins lucrativos",
    "pessoa juridica de direito privado com fins lucrativos sociedade civil": "Privada com fins lucrativos",
    "pessoa juridica de direito privado com fins lucrativos sociedade mercantil ou comercial": "Privada com fins lucrativos",
    "pessoa juridica de direito privado com fins lucrativos associacao de utilidade publica": "Privada com fins lucrativos",
    "pessoa juridica de direito privado sem fins lucrativos": "Privada sem fins lucrativos",
    "pessoa juridica de direito privado sem fins lucrativos fundacao": "Privada sem fins lucrativos",
    "pessoa juridica de direito privado sem fins lucrativos associacao de utilidade publica": "Privada sem fins lucrativos",
    "pessoa juridica de direito privado sem fins lucrativos sociedade": "Privada sem fins lucrativos",
}


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


# Reescreve rótulos do ENADE 2008/2011 'ENGENHARIA (GRUPO X) - <CURSO>' para
# o nome moderno e colapsa o agregado 'ENGENHARIA (GRUPO X)' em 'ENGENHARIA'
# antes da deduplicação, evitando linhas duplicadas para o mesmo código.
ENGENHARIA_GRUPO_RE = re.compile(
    r"^\s*ENGENHARIA\s*\(GRUPO\s+[IVX]+\)\s*(?:-\s*(.+?))?\s*$",
    re.IGNORECASE,
)


def strip_accents(value) -> str:
    s = unicodedata.normalize("NFKD", str(value))
    return "".join(c for c in s if not unicodedata.combining(c))


def canon_key(value) -> str:
    s = strip_accents(value).lower()
    s = re.sub(r"[^\w\s]", " ", s, flags=re.UNICODE)
    return re.sub(r"\s+", " ", s).strip()


def map_canonical(value, table: dict[str, str]):
    if pd.isna(value):
        return value
    s = str(value).strip()
    if not s:
        return None
    return table.get(canon_key(s), s).upper()


def clean_uf(value):
    if pd.isna(value):
        return value
    s = str(value).strip()
    m = re.search(r"\(([A-Za-z]{2})\)\s*$", s)
    if m:
        return m.group(1).upper()
    if len(s) == 2 and s.isalpha():
        return s.upper()
    s_upper = strip_accents(s).upper()
    return UF_FULL_TO_ABBREV.get(s_upper, s.upper())


def title_case_with_connectors(value):
    if pd.isna(value):
        return value
    s = re.sub(r"\s+", " ", str(value)).strip()
    if not s:
        return value
    tokens = re.split(r"([\s\-/])", s)
    out = []
    seen_word = False
    for tok in tokens:
        if not tok:
            continue
        if not re.search(r"\w", tok, flags=re.UNICODE):
            out.append(tok)
            continue
        low = tok.lower()
        if seen_word and low in PORTUGUESE_CONNECTORS:
            out.append(low)
        else:
            out.append(low[0].upper() + low[1:])
        seen_word = True
    return "".join(out)


def strip_double_quotes(value):
    if pd.isna(value):
        return value
    cleaned = str(value).replace('"', "").strip()
    return cleaned if cleaned else None


def drop_engenharia_grupo(value) -> str:
    if pd.isna(value):
        return value
    s = str(value).strip()
    match = ENGENHARIA_GRUPO_RE.match(s)
    if not match:
        return s
    sub = match.group(1)
    return sub.strip() if sub else "ENGENHARIA"
