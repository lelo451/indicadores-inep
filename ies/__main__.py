"""Executa o pipeline completo: censo (1) → final (2) → indicadores (3)."""
from __future__ import annotations

import pandas as pd

from . import censo, final
from ._common import FINAL_XLSX


def main() -> None:
    print("=== Etapa 1: consolidando Censo da Educação Superior ===")
    censo.run()
    print()
    print("=== Etapa 2: mesclando e-MEC com Censo consolidado ===")
    final.run()
    print()
    print("=== Etapa 3: reaplicando metadados em indicadores_consolidados ===")
    final.apply_to_indicadores(pd.read_excel(FINAL_XLSX))


if __name__ == "__main__":
    main()
