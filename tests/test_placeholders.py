from __future__ import annotations

import unittest

from corporate_report_compiler.diagnostics import Diagnostics
from corporate_report_compiler.placeholders import (
    normalize_identifier,
    parse_column_placeholder,
    validate_text_placeholders,
)


class PlaceholderTests(unittest.TestCase):
    def test_identifier_normalizes_case_and_space(self) -> None:
        diagnostics = Diagnostics()
        result = normalize_identifier("  v_dni$1  ", label="Alias", diagnostics=diagnostics)
        self.assertEqual(result, "V_DNI$1")
        self.assertFalse(diagnostics.has_errors)

    def test_identifier_rejects_digit_prefix_unicode_and_overlength(self) -> None:
        for value in ("1DNI", "ÁREA", "A" * 31, "A-B"):
            with self.subTest(value=value):
                diagnostics = Diagnostics()
                self.assertIsNone(normalize_identifier(value, label="Alias", diagnostics=diagnostics))
                self.assertTrue(diagnostics.has_errors)

    def test_header_allows_title_field_and_builtins(self) -> None:
        diagnostics = Diagnostics()
        fields = validate_text_placeholders(
            "{{REPORT_TITLE}}\n{{FIELD:office_name}} — {{APP_USER}} — {{GENERATED_AT}}",
            allowed_kinds={"REPORT_TITLE", "FIELD", "APP_USER", "GENERATED_AT"},
            diagnostics=diagnostics,
            location="Encabezado",
        )
        self.assertEqual(fields, ("OFFICE_NAME",))
        self.assertFalse(diagnostics.has_errors)

    def test_unknown_expression_is_rejected(self) -> None:
        diagnostics = Diagnostics()
        validate_text_placeholders(
            "{{FIELD:DNI | upper}}",
            allowed_kinds={"FIELD"},
            diagnostics=diagnostics,
            location="Encabezado",
        )
        self.assertTrue(diagnostics.has_errors)
        self.assertIn("TOKEN-003", {item.code for item in diagnostics.items})

    def test_unbalanced_braces_are_rejected(self) -> None:
        for text in ("{{FIELD:DNI}", "{FIELD:DNI}}", "{{FIELD:DNI}} }}"):
            with self.subTest(text=text):
                diagnostics = Diagnostics()
                validate_text_placeholders(
                    text,
                    allowed_kinds={"FIELD"},
                    diagnostics=diagnostics,
                    location="Encabezado",
                )
                self.assertTrue(diagnostics.has_errors)

    def test_column_marker_is_whole_cell_case_insensitive(self) -> None:
        diagnostics = Diagnostics()
        self.assertEqual(
            parse_column_placeholder("  {{ column : vnom }}  ", diagnostics, "cell"),
            "VNOM",
        )
        self.assertFalse(diagnostics.has_errors)

    def test_column_marker_rejects_extra_text_and_expressions(self) -> None:
        for text in ("Nombre {{COLUMN:VNOM}}", "{{COLUMN:VNOM|upper}}", "{{FIELD:VNOM}}"):
            with self.subTest(text=text):
                diagnostics = Diagnostics()
                self.assertIsNone(parse_column_placeholder(text, diagnostics, "cell"))
                self.assertTrue(diagnostics.has_errors)


if __name__ == "__main__":
    unittest.main()

