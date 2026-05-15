"""Executa os dois consolidadores de dicionários em sequência: ENADE e IDD."""
from __future__ import annotations

from . import enade, idd


def main() -> None:
    print("=== Dicionário ENADE ===")
    enade.main()
    print()
    print("=== Dicionário IDD ===")
    idd.main()


if __name__ == "__main__":
    main()
