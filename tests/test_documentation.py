"""Los entregables Word contienen las instrucciones críticas del contrato."""

from __future__ import annotations

import unittest
from pathlib import Path

from docx import Document


ROOT = Path(__file__).resolve().parents[1]
REQUIRED = {
    "deliverables/Manual_de_uso.docx": (
        "Reload on Submit",
        "Always",
        "Submit Page",
        "When Button Pressed",
        "P0_REPORT_FORMAT",
        "Compilador de reportes APEX",
        "Nuevo proyecto desde DOCX",
        "SQL-013",
    ),
    "deliverables/Informe_de_alcance_y_limitaciones.docx": ("chr(38)", "bytes UTF-8"),
    "docs/CONTRACT.docx": ("bytes UTF-8", "ecuaciones", "nombres reservados", "chr(38)"),
    "docs/RELEASE_REVIEW.docx": ("GO para empaquetado", "Reload on Submit"),
    "docs/TEMPLATE_QA.docx": ("templates",),
    "README.docx": ("Abrir compilador.bat",),
    "sql/README.docx": ("Reload on Submit", "-20173"),
    "examples/README.docx": ("build\\entidades",),
}


def document_text(path: Path) -> str:
    document = Document(str(path))
    parts = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            parts.extend(cell.text for cell in row.cells)
    return "\n".join(parts)


class DeliverableDocumentationTests(unittest.TestCase):
    def test_word_documents_contain_critical_instructions(self) -> None:
        for relative, phrases in REQUIRED.items():
            with self.subTest(document=relative):
                path = ROOT / relative
                self.assertTrue(path.is_file(), f"Falta {relative}")
                text = document_text(path)
                for phrase in phrases:
                    self.assertIn(phrase, text, f"{relative} no contiene {phrase!r}")


if __name__ == "__main__":
    unittest.main()
