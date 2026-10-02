"""Abre la interfaz gráfica del compilador sin instalar el paquete.

Lo usa «Abrir compilador.bat» (y el acceso directo de la carpeta raíz). Solo
requiere Python 3.12+ con python-docx, que el .bat prepara en .venv.
"""

from pathlib import Path
import sys


PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from corporate_report_compiler.gui import main  # noqa: E402


if __name__ == "__main__":
    raise SystemExit(main())
