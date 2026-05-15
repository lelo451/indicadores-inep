"""Executa o pipeline completo: censo (etapa 1) → final (etapa 2)."""
from __future__ import annotations

from . import censo, final


def main() -> None:
    print("=== Etapa 1: consolidando Censo da Educação Superior ===")
    censo.run()
    print()
    print("=== Etapa 2: mesclando e-MEC com Censo consolidado ===")
    final.run()


if __name__ == "__main__":
    main()
