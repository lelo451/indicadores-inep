# -*- coding: utf-8 -*-
"""Consolida o 'DICIONÁRIO' dos microdados do IDD.

CO_UF_CURSO só aparece em 2014; CO_MODALIDADE recebe '1 = Presencial' nos
anos que não trazem essa variável. Aliases usados em edições antigas
(co_catad, co_orgac, co_munic) são mapeados aos nomes canônicos.
"""

from __future__ import annotations

import os

from dicionario_comum import consolidate

DATA_ROOT = os.path.dirname(os.path.abspath(__file__))


def main() -> None:
    consolidate(
        dict_dir=os.path.join(DATA_ROOT, "Microdados", "idd"),
        dict_pattern="dicionarios idd *.xls*",
        variables=[
            "CO_CATEGAD",
            "CO_ORGACAD",
            "CO_GRUPO",
            "CO_MODALIDADE",
            "CO_UF_CURSO",
            "CO_MUNIC_CURSO",
        ],
        aliases={
            "CO_CATAD": "CO_CATEGAD",
            "CO_ORGAC": "CO_ORGACAD",
            "CO_MUNIC": "CO_MUNIC_CURSO",
        },
        municipios_csv=os.path.join(DATA_ROOT, "Microdados", "municipios.csv"),
        output_path=os.path.join(DATA_ROOT, "dicionario_idd_consolidado.xlsx"),
    )


if __name__ == "__main__":
    main()
