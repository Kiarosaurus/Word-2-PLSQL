from __future__ import annotations

import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PACKAGE_PATH = ROOT / "sql" / "modo_simple" / "pkg_corporate_reports.sql"


def package_spec(source: str) -> str:
    match = re.search(
        r"CREATE\s+OR\s+REPLACE\s+PACKAGE\s+pkg_corporate_reports"
        r"(?:\s+AUTHID\s+(?:CURRENT_USER|DEFINER))?\s+AS"
        r"(?P<spec>.*?)END\s+pkg_corporate_reports\s*;\s*/",
        source,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if match is None:
        raise AssertionError("No se encontró la especificación del package.")
    return match.group("spec")


def package_body(source: str) -> str:
    match = re.search(
        r"CREATE\s+OR\s+REPLACE\s+PACKAGE\s+BODY\s+pkg_corporate_reports\s+AS"
        r"(?P<body>.*?)END\s+pkg_corporate_reports\s*;\s*/",
        source,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if match is None:
        raise AssertionError("No se encontró el cuerpo del package.")
    return match.group("body")


def procedure_body(source: str, procedure: str) -> str:
    """Return one top-level package procedure implementation.

    The terminating name is part of the package coding convention and avoids
    accidentally satisfying a contract with code from the next procedure.
    """

    match = re.search(
        rf"\bPROCEDURE\s+{re.escape(procedure)}\s*\(.*?"
        rf"\bEND\s+{re.escape(procedure)}\s*;",
        source,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if match is None:
        raise AssertionError(f"No se encontró la implementación de {procedure}.")
    return match.group(0)


class PackageStaticContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.source = PACKAGE_PATH.read_text(encoding="utf-8")
        cls.spec = package_spec(cls.source)
        cls.body = package_body(cls.source)

    def assert_signature_contains(self, procedure: str, required_fragments: tuple[str, ...]) -> None:
        match = re.search(
            rf"PROCEDURE\s+{procedure}\s*\((?P<args>.*?)\)\s*;",
            self.spec,
            flags=re.IGNORECASE | re.DOTALL,
        )
        self.assertIsNotNone(match, f"Falta la API pública {procedure}.")
        args = re.sub(r"\s+", " ", match.group("args")).upper()  # type: ignore[union-attr]
        for fragment in required_fragments:
            self.assertIn(re.sub(r"\s+", " ", fragment).upper(), args)

    def test_v7_public_ig_api_remains_frozen(self) -> None:
        common = (
            "p_application_id IN NUMBER",
            "p_page_id IN NUMBER",
            "p_region_static_id IN VARCHAR2",
            "p_visible_columns_json IN CLOB",
            "p_title IN VARCHAR2",
            "p_file_name IN VARCHAR2",
            "p_generated_by IN VARCHAR2 DEFAULT NULL",
            "p_orientation IN VARCHAR2 DEFAULT 'AUTO'",
            "p_max_rows IN PLS_INTEGER DEFAULT 1000",
            "p_excluded_columns_json IN CLOB DEFAULT NULL",
            "p_column_spans_json IN CLOB DEFAULT NULL",
            "p_display_columns_json IN CLOB DEFAULT NULL",
        )
        self.assert_signature_contains("download_ig", common + ("p_format IN VARCHAR2 DEFAULT 'PDF'",))
        self.assert_signature_contains("download_ig_pdf", common)

    def test_query_api_matches_compiler_contract(self) -> None:
        self.assert_signature_contains(
            "download_query",
            (
                "p_sql_query IN VARCHAR2",
                "p_columns_json IN CLOB",
                "p_bindings_json IN CLOB DEFAULT NULL",
                "p_fields_json IN CLOB DEFAULT NULL",
                "p_style_json IN CLOB DEFAULT NULL",
                "p_header_template IN VARCHAR2 DEFAULT '{{REPORT_TITLE}}'",
                "p_footer_template IN VARCHAR2 DEFAULT NULL",
                "p_title IN VARCHAR2 DEFAULT 'Reporte'",
                "p_file_name IN VARCHAR2 DEFAULT 'reporte'",
                "p_format IN VARCHAR2 DEFAULT 'PDF'",
                "p_orientation IN VARCHAR2 DEFAULT 'AUTO'",
                "p_excluded_columns_json IN CLOB DEFAULT NULL",
                "p_column_widths_json IN CLOB DEFAULT NULL",
            ),
        )

    def test_query_api_does_not_expose_internal_layout_parameters(self) -> None:
        signature = re.search(
            r"PROCEDURE\s+download_query\s*\((?P<args>.*?)\)\s*;",
            self.spec,
            flags=re.IGNORECASE | re.DOTALL,
        )
        self.assertIsNotNone(signature)
        args = signature.group("args").lower()  # type: ignore[union-attr]
        self.assertNotIn("p_paper_size", args)
        self.assertNotIn("p_auto_reserve_pct", args)

        implementation = procedure_body(self.body, "download_query").lower()
        body_signature = implementation.split("\n    is", 1)[0]
        self.assertNotIn("p_paper_size", body_signature)
        self.assertNotIn("p_auto_reserve_pct", body_signature)
        self.assertIn("apex_data_export.c_size_a4", implementation)
        self.assertIn("c_auto_reserve_pct", implementation)

    def test_package_keeps_invoker_rights(self) -> None:
        self.assertRegex(
            self.source,
            r"(?i)CREATE\s+OR\s+REPLACE\s+PACKAGE\s+"
            r"pkg_corporate_reports\s+AUTHID\s+CURRENT_USER\s+AS",
        )

    def test_query_runtime_uses_bound_apex_exec_context(self) -> None:
        lowered = self.source.lower()
        self.assertIn("apex_exec.open_query_context", lowered)
        self.assertRegex(lowered, r"p_auto_bind_items\s*=>\s*false")
        self.assertIn("apex_exec.add_parameter", lowered)
        self.assertIn("apex_exec.get_column_position", lowered)
        self.assertNotRegex(lowered, r"execute\s+immediate\s+p_sql_query")

    def test_field_usage_check_does_not_embed_identifier_in_regex(self) -> None:
        lowered = self.source.lower()
        self.assertIn("function template_uses_field", lowered)
        self.assertIn("template_uses_field(l_template_text, l_key)", lowered)
        self.assertNotIn("'[[:space:]]*' || l_key", lowered)

    def test_field_runtime_accepts_same_optional_spaces_as_compiler(self) -> None:
        self.assertIn("REGEXP_LIKE(l_token, '^FIELD[[:space:]]*:'", self.source)
        self.assertIn("'^FIELD[[:space:]]*:[[:space:]]*'", self.source)

    def test_context_is_closed_on_success_and_exception(self) -> None:
        body = procedure_body(self.body, "download_query").lower()
        success_close = body.find("apex_exec.close(l_context)")
        download = body.find("apex_data_export.download")
        self.assertGreaterEqual(success_close, 0, "Falta cerrar el contexto en la ruta exitosa.")
        self.assertGreater(download, success_close, "El contexto debe cerrarse antes de iniciar la descarga.")

        # DOWNLOAD_QUERY contiene bloques BEGIN/EXCEPTION internos para
        # conversiones tipadas. El manejador relevante es el EXCEPTION de
        # nivel superior, indentado al mismo nivel que el BEGIN principal.
        top_level_exception = body.rfind("\n    exception")
        self.assertGreaterEqual(top_level_exception, 0, "Falta el bloque EXCEPTION principal.")
        exception = re.search(
            r"\bwhen\s+others\s+then\b(?P<handler>.*?)\braise\s*;",
            body[top_level_exception:],
            flags=re.DOTALL,
        )
        self.assertIsNotNone(exception, "Falta el manejador WHEN OTHERS de DOWNLOAD_QUERY.")
        handler = exception.group("handler")  # type: ignore[union-attr]
        self.assertRegex(
            handler,
            r"if\s+l_context_is_open\s+then\s+apex_exec\.close\(l_context\)\s*;\s+end\s+if\s*;",
        )

    def test_package_uses_expected_native_export_api(self) -> None:
        lowered = self.source.lower()
        self.assertIn("apex_data_export.export", lowered)
        self.assertIn("apex_data_export.download", lowered)
        self.assertIn("apex_data_export.add_column", lowered)

    def test_script_has_both_compilation_units_and_sqlplus_delimiters(self) -> None:
        self.assertRegex(
            self.source,
            r"(?i)CREATE\s+OR\s+REPLACE\s+PACKAGE\s+pkg_corporate_reports"
            r"(?:\s+AUTHID\s+(?:CURRENT_USER|DEFINER))?\s+AS",
        )
        self.assertRegex(self.source, r"(?i)CREATE\s+OR\s+REPLACE\s+PACKAGE\s+BODY\s+pkg_corporate_reports\s+AS")
        self.assertGreaterEqual(len(re.findall(r"(?m)^/\s*$", self.source)), 2)


if __name__ == "__main__":
    unittest.main()
