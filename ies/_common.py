"""Caminhos e helpers compartilhados pelo pipeline."""
from __future__ import annotations

from pathlib import Path

import pandas as pd


# O pacote vive em ``<projeto>/ies/``; ROOT é a raiz do projeto.
ROOT = Path(__file__).resolve().parent.parent

# Entradas
SRC_DIR = ROOT / "Lista INEP Censo Superior"
WHITELIST_CSV = ROOT / "list_ies.csv"
BASE_CSV = ROOT / "Microdados" / "ies_siglas.csv"

# Saídas do pipeline (todas as planilhas geradas vão para ``outputs/``).
OUTPUTS_DIR = ROOT / "outputs"
CONSOLIDADA_XLSX = OUTPUTS_DIR / "lista_ies_consolidada.xlsx"
FINAL_XLSX = OUTPUTS_DIR / "list_ies_final.xlsx"
# Produzido por ``consolidar_indicadores.py``; usado como 3º nível de
# suplemento por ``ies.final`` para IES que e-MEC e Censo não cobrem.
INDICADORES_XLSX = OUTPUTS_DIR / "indicadores_consolidados.xlsx"


def normalize_code(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").astype("Int64")


def clean_text(series: pd.Series) -> pd.Series:
    return series.astype("string").str.strip()
