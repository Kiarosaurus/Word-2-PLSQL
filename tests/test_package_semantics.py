"""Contrato estático del package con comentarios y literales excluidos.

Las pruebas de ``test_package_static`` buscan texto sobre el fuente completo y
podían satisfacerse con código comentado o con una subcadena de otra rutina.
Estas comprobaciones operan sobre el código efectivo de cada procedimiento y
se validan a sí mismas contra mutaciones semánticas conocidas.
"""

from __future__ import annotations

import re
import unittest
from pathlib import Path

from tests.test_package_static import PACKAGE_PATH, package_body, package_spec


DOWNLOAD_QUERY_SIGNATURE = (
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
    "p_max_rows IN PLS_INTEGER DEFAULT 1000",
    "p_excluded_columns_json IN CLOB DEFAULT NULL",
    "p_column_widths_json IN CLOB DEFAULT NULL",
)
IG_COMMON = (
    "p_application_id IN NUMBER",
    "p_page_id IN NUMBER",
    "p_region_static_id IN VARCHAR2",
    "p_visible_columns_json IN CLOB",
    "p_title IN VARCHAR2",
    "p_file_name IN VARCHAR2",
    "p_generated_by IN VARCHAR2 DEFAULT NULL",
)
IG_TAIL = (
    "p_orientation IN VARCHAR2 DEFAULT 'AUTO'",
    "p_max_rows IN PLS_INTEGER DEFAULT 1000",
    "p_excluded_columns_json IN CLOB DEFAULT NULL",
    "p_column_spans_json IN CLOB DEFAULT NULL",
    "p_display_columns_json IN CLOB DEFAULT NULL",
)
EXPECTED_SIGNATURES = {
    "download_query": DOWNLOAD_QUERY_SIGNATURE,
    "download_ig": IG_COMMON + ("p_format IN VARCHAR2 DEFAULT 'PDF'",) + IG_TAIL,
    "download_ig_pdf": IG_COMMON + IG_TAIL,
}


def strip_plsql_comments(source: str) -> str:
    """Sustituye comentarios ``--`` y ``/* */`` por espacios, respetando literales."""

    result: list[str] = []
    index = 0
    length = len(source)
    while index < length:
        char = source[index]
        pair = source[index : index + 2]
        if pair == "--":
            end = source.find("\n", index)
            end = length if end < 0 else end
            result.append(" " * (end - index))
            index = end
        elif pair == "/*":
            end = source.find("*/", index + 2)
            end = length if end < 0 else end + 2
            result.append(re.sub(r"[^\n]", " ", source[index:end]))
            index = end
        elif source[index : index + 3].lower() == "q'~":
            end = source.find("~'", index + 3)
            end = length if end < 0 else end + 2
            result.append(source[index:end])
            index = end
        elif char == "'":
            end = index + 1
            while end < length:
                if source[end] == "'" and source[end + 1 : end + 2] == "'":
                    end += 2
                elif source[end] == "'":
                    end += 1
                    break
                else:
                    end += 1
            result.append(source[index:end])
            index = end
        else:
            result.append(char)
            index += 1
    return "".join(result)


def unit(source: str, kind: str, name: str) -> str:
    match = re.search(
        rf"\b{kind}\s+{re.escape(name)}\b.*?\bEND\s+{re.escape(name)}\s*;",
        source,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if match is None:
        raise AssertionError(f"No se encontró {kind} {name}.")
    return match.group(0)


def spec_signature(source: str, procedure: str) -> tuple[str, ...]:
    spec = strip_plsql_comments(package_spec(source))
    match = re.search(
        rf"PROCEDURE\s+{procedure}\s*\((?P<args>.*?)\)\s*;",
        spec,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if match is None:
        raise AssertionError(f"Falta la API pública {procedure}.")
    return tuple(re.sub(r"\s+", " ", part.strip()) for part in match.group("args").split(","))


def check_contract(source: str) -> None:
    """Lanza AssertionError si el contrato efectivo del package no se cumple."""

    for procedure, expected in EXPECTED_SIGNATURES.items():
        actual = spec_signature(source, procedure)
        if tuple(item.upper() for item in actual) != tuple(item.upper() for item in expected):
            raise AssertionError(f"La firma de {procedure} cambió: {actual}")

    if not re.search(
        r"(?i)CREATE\s+OR\s+REPLACE\s+PACKAGE\s+pkg_corporate_reports\s+AUTHID\s+CURRENT_USER\s+AS",
        strip_plsql_comments(source),
    ):
        raise AssertionError("El package debe usar AUTHID CURRENT_USER.")

    body = strip_plsql_comments(package_body(source))
    query = unit(body, "PROCEDURE", "download_query").lower()

    if not re.search(r"apex_exec\.open_query_context\s*\((?:(?!\);).)*p_auto_bind_items\s*=>\s*false", query, re.S):
        raise AssertionError("OPEN_QUERY_CONTEXT debe usar p_auto_bind_items => FALSE.")
    if re.search(r"p_auto_bind_items\s*=>\s*true", query):
        raise AssertionError("DOWNLOAD_QUERY no puede activar el auto-bind.")
    if re.search(r"\bexecute\s+immediate\b|\bdbms_sql\b", query):
        raise AssertionError("DOWNLOAD_QUERY no puede ejecutar SQL dinámico propio.")
    if not re.search(r"p_sql_query\s*=>\s*p_sql_query", query):
        raise AssertionError("El SQL debe pasarse sin modificar a APEX_EXEC.")

    exception_at = query.rfind("\n    exception")
    success, handlers = query[:exception_at], query[exception_at:]
    close_at = success.rfind("apex_exec.close(l_context)")
    if close_at < 0 or success.find("apex_data_export.download", close_at) < 0:
        raise AssertionError("El contexto debe cerrarse antes de DOWNLOAD en la ruta exitosa.")
    for handler in ("apex_application.e_stop_apex_engine", "others"):
        match = re.search(rf"when\s+{re.escape(handler)}\s+then(?P<code>.*?)\braise\s*;", handlers, re.S)
        if match is None or "apex_exec.close(l_context)" not in match.group("code"):
            raise AssertionError(f"El manejador {handler} debe cerrar el contexto.")

    guard = unit(body, "PROCEDURE", "assert_read_only_query").lower()
    if re.search(r"ltrim\s*\([^,()]*(?:\([^()]*\))?[^,()]*\)", guard):
        raise AssertionError("LTRIM sin conjunto solo elimina espacios; use c_blank.")
    if "chr(9)" not in guard or "chr(10)" not in guard or "chr(13)" not in guard:
        raise AssertionError("El guard SQL debe eliminar tabulador, LF y CR como el compilador.")

    resolver = unit(body, "FUNCTION", "resolve_text_template").lower()
    if re.search(r"\blength\s*\(", resolver) or "lengthb(" not in resolver:
        raise AssertionError("Encabezado y pie deben limitarse en bytes.")

    if "'app_user'" not in query or "-20173" not in query:
        raise AssertionError("Los nombres reservados de APEX deben rechazarse como binds.")
    if re.search(r"if\s+sqlcode\s+between\s+-20999\s+and\s+-20000", query):
        raise AssertionError("Los errores de APEX deben sanearse, no propagarse.")


class PackageSemanticContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.source = PACKAGE_PATH.read_text(encoding="utf-8")

    def test_effective_contract_holds(self) -> None:
        check_contract(self.source)

    def test_contract_rejects_semantic_mutations(self) -> None:
        mutations = {
            "auto-bind comentado": (
                "p_auto_bind_items => FALSE",
                "p_auto_bind_items => TRUE -- p_auto_bind_items => FALSE",
            ),
            "cierre exitoso comentado": (
                "        apex_exec.close(l_context);\n        l_context_is_open := FALSE;\n\n        apex_data_export.download(",
                "        NULL; -- apex_exec.close(l_context);\n        l_context_is_open := FALSE;\n\n        apex_data_export.download(",
            ),
            "default ampliado": (
                "p_max_rows               IN PLS_INTEGER DEFAULT 1000,\n        p_excluded_columns_json  IN CLOB DEFAULT NULL,\n        p_column_widths_json",
                "p_max_rows               IN PLS_INTEGER DEFAULT 100000,\n        p_excluded_columns_json  IN CLOB DEFAULT NULL,\n        p_column_widths_json",
            ),
            "parámetros intercambiados": (
                "        p_title                  IN VARCHAR2 DEFAULT 'Reporte',\n        p_file_name              IN VARCHAR2 DEFAULT 'reporte',\n        p_format                 IN VARCHAR2 DEFAULT 'PDF',\n        p_orientation            IN VARCHAR2 DEFAULT 'AUTO',\n        p_max_rows               IN PLS_INTEGER DEFAULT 1000,\n        p_excluded_columns_json  IN CLOB DEFAULT NULL,\n        p_column_widths_json     IN CLOB DEFAULT NULL\n    );\n\nEND",
                "        p_file_name              IN VARCHAR2 DEFAULT 'reporte',\n        p_title                  IN VARCHAR2 DEFAULT 'Reporte',\n        p_format                 IN VARCHAR2 DEFAULT 'PDF',\n        p_orientation            IN VARCHAR2 DEFAULT 'AUTO',\n        p_max_rows               IN PLS_INTEGER DEFAULT 1000,\n        p_excluded_columns_json  IN CLOB DEFAULT NULL,\n        p_column_widths_json     IN CLOB DEFAULT NULL\n    );\n\nEND",
            ),
            "cierre en stop engine eliminado": (
                "        WHEN apex_application.e_stop_apex_engine THEN\n            IF l_context_is_open THEN\n                apex_exec.close(l_context);",
                "        WHEN apex_application.e_stop_apex_engine THEN\n            IF l_context_is_open THEN\n                NULL;",
            ),
            "SQL dinámico": (
                "        l_context := apex_exec.open_query_context(",
                "        EXECUTE IMMEDIATE 'BEGIN NULL; END;';\n        l_context := apex_exec.open_query_context(",
            ),
            "LTRIM de un argumento": (
                "l_check := LTRIM(p_sql_query, c_blank);",
                "l_check := LTRIM(p_sql_query);",
            ),
        }
        for label, (before, after) in mutations.items():
            with self.subTest(mutation=label):
                # DOWNLOAD_QUERY es la última rutina: se muta la última
                # aparición para no tocar el flujo IG, que comparte fragmentos.
                self.assertGreaterEqual(self.source.count(before), 1, f"Falta el ancla de {label}.")
                head, _separator, tail = self.source.rpartition(before)
                mutated = head + after + tail
                with self.assertRaises(AssertionError):
                    check_contract(mutated)

    def test_comment_stripper_keeps_literals(self) -> None:
        sample = "x := 'a -- b' || q'~/* c */~'; -- fin\n/* bloque */ y"
        stripped = strip_plsql_comments(sample)
        self.assertIn("'a -- b'", stripped)
        self.assertIn("q'~/* c */~'", stripped)
        self.assertNotIn("fin", stripped)
        self.assertNotIn("bloque", stripped)


if __name__ == "__main__":
    unittest.main()
