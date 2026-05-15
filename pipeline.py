"""Roda o pipeline completo de indicadores na ordem correta.

Entrada recomendada para regerar todos os artefatos. Garante que
``consolidar_indicadores`` rode antes do pacote ``ies/``, evitando que a etapa
3 de ``ies.final`` (que reaplica metadados em ``indicadores_consolidados``)
seja sobrescrita por uma execução tardia do consolidador.

Pré-requisitos manuais (não rodados aqui, mudam raramente):
- ``python -m dicionario``      → outputs/dicionario_{enade,idd}_consolidado.xlsx
- ``python enrich_ies_sigla.py`` → Microdados/ies_siglas.csv
"""
from __future__ import annotations

import consolidar_indicadores
from ies import __main__ as ies_main


def main() -> None:
    print("### Stage A: consolidar_indicadores ###")
    consolidar_indicadores.main()
    print()
    print("### Stage B: ies (censo + final + apply) ###")
    ies_main.main()


if __name__ == "__main__":
    main()
