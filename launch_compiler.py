"""Iniciador portable del compilador desde el directorio del proyecto."""

from pathlib import Path
import sys


PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from corporate_report_compiler.cli import main  # noqa: E402


if __name__ == "__main__":
    raise SystemExit(main())

