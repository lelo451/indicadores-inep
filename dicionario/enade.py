# -*- coding: utf-8 -*-
"""Consolida o 'DICIONÁRIO DE VARIÁVEIS' dos microdados do ENADE."""

from __future__ import annotations

import os

from . import consolidate

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main() -> None:
    consolidate(
        dict_dir=os.path.join(ROOT, "Microdados", "enade"),
        dict_pattern="dicionarios enade *.xls*",
        variables=[
            "CO_CATEGAD",
            "CO_ORGACAD",
            "CO_GRUPO",
            "CO_MODALIDADE",
            "CO_UF_CURSO",
            "CO_MUNIC_CURSO",
        ],
        municipios_csv=os.path.join(ROOT, "Microdados", "municipios.csv"),
        output_path=os.path.join(ROOT, "outputs", "dicionario_enade_consolidado.xlsx"),
    )


if __name__ == "__main__":
    main()
