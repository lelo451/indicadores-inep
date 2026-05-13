# -*- coding: utf-8 -*-
"""Consolida o 'DICIONÁRIO DE VARIÁVEIS' dos microdados do ENADE."""

from __future__ import annotations

import os

from dicionario_comum import consolidate

DATA_ROOT = os.path.dirname(os.path.abspath(__file__))


def main() -> None:
    consolidate(
        dict_dir=os.path.join(DATA_ROOT, "Microdados", "enade"),
        dict_pattern="dicionarios enade *.xls*",
        variables=[
            "CO_CATEGAD",
            "CO_ORGACAD",
            "CO_GRUPO",
            "CO_MODALIDADE",
            "CO_UF_CURSO",
            "CO_MUNIC_CURSO",
        ],
        municipios_csv=os.path.join(DATA_ROOT, "Microdados", "municipios.csv"),
        output_path=os.path.join(DATA_ROOT, "dicionario_enade_consolidado.xlsx"),
    )


if __name__ == "__main__":
    main()
