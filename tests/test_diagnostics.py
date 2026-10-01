from __future__ import annotations

import unittest

from corporate_report_compiler.diagnostics import Diagnostic, Diagnostics


class DiagnosticsTests(unittest.TestCase):
    def test_strict_warning_can_be_promoted_to_error(self) -> None:
        diagnostics = Diagnostics(strict=True)
        diagnostics.warning("W-002", "warning", strict_error=True)
        self.assertTrue(diagnostics.has_errors)
        self.assertEqual(diagnostics.error_count, 1)
        self.assertEqual(diagnostics.warning_count, 0)

    def test_compatible_warning_remains_warning(self) -> None:
        diagnostics = Diagnostics(strict=False)
        diagnostics.warning("W-002", "warning", strict_error=True)
        self.assertFalse(diagnostics.has_errors)
        self.assertEqual(diagnostics.warning_count, 1)

    def test_serialization_is_sorted_and_omits_none(self) -> None:
        diagnostics = Diagnostics()
        diagnostics.extend(
            [
                Diagnostic("W-002", "WARNING", "second"),
                Diagnostic("E-002", "ERROR", "error b", "z"),
                Diagnostic("E-001", "ERROR", "error a"),
                Diagnostic("W-001", "WARNING", "first", suggestion="fix"),
            ]
        )
        payload = diagnostics.to_dict()
        self.assertEqual([item["code"] for item in payload["diagnostics"]], ["E-001", "E-002", "W-001", "W-002"])
        self.assertNotIn("location", payload["diagnostics"][0])
        self.assertEqual(payload["errors"], 2)
        self.assertEqual(payload["warnings"], 2)


if __name__ == "__main__":
    unittest.main()

