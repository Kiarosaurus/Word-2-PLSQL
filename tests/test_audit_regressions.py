"""Regresiones de los hallazgos confirmados en la auditoría 1.0.0.

Cada prueba reproduce un caso que antes se aceptaba localmente y fallaba o
cambiaba en ejecución, o que provocaba una excepción no controlada.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from corporate_report_compiler import compiler as compiler_module
from corporate_report_compiler.compiler import compile_project, validate_project
from corporate_report_compiler.diagnostics import Diagnostics
from corporate_report_compiler.emitter import oracle_expression
from corporate_report_compiler.placeholders import validate_text_placeholders
from corporate_report_compiler.security import inspect_docx_zip
from corporate_report_compiler.sql_reader import validate_sql

from tests.support import build_valid_docx, make_valid_project, rewrite_zip


ROOT = Path(__file__).resolve().parents[1]


def codes(diagnostics: Diagnostics) -> set[str]:
    return {item.code for item in diagnostics.items}


def result_codes(result) -> set[str]:
    return {item.code for item in result.diagnostics.items}


class TemporaryTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def project(self, **kwargs) -> Path:
        return make_valid_project(self.root / "project", **kwargs)

    def rewrite_project(self, project_path: Path, **changes: object) -> None:
        payload = json.loads(project_path.read_text(encoding="utf-8"))
        payload.update(changes)
        project_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

    def mutate_document(self, project_path: Path, transform) -> None:
        template = project_path.parent / "template.docx"
        with zipfile.ZipFile(template) as archive:
            xml = archive.read("word/document.xml").decode("utf-8")
        mutated = transform(xml)
        self.assertNotEqual(mutated, xml, "La mutación no cambió el documento.")
        hostile = project_path.parent / "hostile.docx"
        rewrite_zip(template, hostile, replacements={"word/document.xml": mutated.encode("utf-8")})
        os.replace(hostile, template)


class SecurityRegressionTests(TemporaryTestCase):
    def setUp(self) -> None:
        super().setUp()
        self.valid = build_valid_docx(self.root / "valid.docx")

    def inspect(self, path: Path) -> Diagnostics:
        diagnostics = Diagnostics()
        inspect_docx_zip(path, diagnostics)
        return diagnostics

    def test_doctype_in_utf16_part_is_rejected(self) -> None:
        part = '<?xml version="1.0" encoding="UTF-16"?><!DOCTYPE x [<!ENTITY e "v">]><x>&e;</x>'
        hostile = rewrite_zip(
            self.valid,
            self.root / "utf16.docx",
            additions={"word/extra.xml": part.encode("utf-16")},
        )
        self.assertIn("OOXML-002", codes(self.inspect(hostile)))

    def test_macro_content_type_is_detected_with_uppercase_part_name(self) -> None:
        with zipfile.ZipFile(self.valid) as source:
            content_types = source.read("[Content_Types].xml").replace(
                b"</Types>",
                b'<Override PartName="/word/x.bin" ContentType="application/vnd.ms-word.macroEnabled.main+xml"/></Types>',
            )
            hostile = self.root / "upper.docx"
            with zipfile.ZipFile(hostile, "w", zipfile.ZIP_DEFLATED) as target:
                for info in source.infolist():
                    if info.filename == "[Content_Types].xml":
                        target.writestr("[CONTENT_TYPES].XML", content_types)
                    else:
                        target.writestr(info, source.read(info.filename))
        self.assertIn("OOXML-005", codes(self.inspect(hostile)))

    def test_corrupt_deflate_stream_is_a_diagnostic_not_an_exception(self) -> None:
        hostile = self.root / "corrupt.docx"
        with zipfile.ZipFile(self.valid) as source:
            with zipfile.ZipFile(hostile, "w", zipfile.ZIP_DEFLATED) as target:
                for info in source.infolist():
                    target.writestr(info, source.read(info.filename))
        data = bytearray(hostile.read_bytes())
        with zipfile.ZipFile(hostile) as archive:
            info = archive.getinfo("word/fontTable.xml")
        start = info.header_offset + 30 + len(info.filename.encode()) + len(info.extra)
        for offset in range(start, start + min(info.compress_size, 40)):
            data[offset] ^= 0xFF
        hostile.write_bytes(bytes(data))
        diagnostics = self.inspect(hostile)
        self.assertTrue(diagnostics.has_errors)

    def test_main_document_part_must_be_word_document_xml(self) -> None:
        with zipfile.ZipFile(self.valid) as source:
            rels = source.read("_rels/.rels").replace(b'Target="word/document.xml"', b'Target="word/main.xml"')
        hostile = rewrite_zip(self.valid, self.root / "renamed.docx", replacements={"_rels/.rels": rels})
        self.assertIn("OOXML-007", codes(self.inspect(hostile)))

    def test_equivalent_main_part_target_is_accepted(self) -> None:
        with zipfile.ZipFile(self.valid) as source:
            rels = source.read("_rels/.rels").replace(b'Target="word/document.xml"', b'Target="./word/document.xml"')
        candidate = rewrite_zip(self.valid, self.root / "dot.docx", replacements={"_rels/.rels": rels})
        self.assertNotIn("OOXML-007", codes(self.inspect(candidate)))

    def test_prohibited_relationship_type_is_rejected_when_part_is_renamed(self) -> None:
        with zipfile.ZipFile(self.valid) as source:
            rels = source.read("word/_rels/document.xml.rels").replace(
                b"</Relationships>",
                b'<Relationship Id="rIdX" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                b'relationships/image" Target="graficos/logo.png"/></Relationships>',
            )
        hostile = rewrite_zip(
            self.valid,
            self.root / "image.docx",
            replacements={"word/_rels/document.xml.rels": rels},
            additions={"word/graficos/logo.png": b"\x89PNG"},
        )
        self.assertIn("OOXML-001", codes(self.inspect(hostile)))


class DocxContentRegressionTests(TemporaryTestCase):
    def assert_rejected(self, transform, expected: str = "OOXML-020") -> None:
        project_path = self.project()
        self.mutate_document(project_path, transform)
        result = validate_project(project_path)
        self.assertFalse(result.valid)
        self.assertIn(expected, result_codes(result))

    def test_inline_wrappers_hiding_text_are_rejected(self) -> None:
        wrappers = {
            "smartTag": '<w:smartTag w:uri="x" w:element="y"><w:r><w:t>X {{BAD}}</w:t></w:r></w:smartTag>',
            "customXml": '<w:customXml w:element="y"><w:r><w:t>X {{BAD}}</w:t></w:r></w:customXml>',
            "dir": '<w:dir w:val="rtl"><w:r><w:t>X {{BAD}}</w:t></w:r></w:dir>',
        }
        for name, fragment in wrappers.items():
            with self.subTest(wrapper=name):
                self.assert_rejected(lambda xml, f=fragment: xml.replace("</w:p>", f + "</w:p>", 1))

    def test_hidden_text_is_rejected(self) -> None:
        self.assert_rejected(
            lambda xml: xml.replace("</w:p>", '<w:r><w:rPr><w:vanish/></w:rPr><w:t>{{APP_USER}}</w:t></w:r></w:p>', 1)
        )

    def test_formatting_revisions_and_comment_markers_are_rejected(self) -> None:
        fragments = (
            '<w:r><w:rPr><w:rPrChange w:id="1" w:author="a"><w:rPr/></w:rPrChange></w:rPr><w:t>x</w:t></w:r>',
            '<w:commentRangeStart w:id="0"/><w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="0"/>',
        )
        for fragment in fragments:
            with self.subTest(fragment=fragment[:30]):
                self.assert_rejected(lambda xml, f=fragment: xml.replace("</w:p>", f + "</w:p>", 1))

    def test_block_level_wrapper_is_rejected(self) -> None:
        self.assert_rejected(
            lambda xml: xml.replace(
                "<w:sectPr",
                '<w:customXml w:element="b"><w:p><w:r><w:t>Libre</w:t></w:r></w:p></w:customXml><w:sectPr',
                1,
            ),
            expected="OOXML-020",
        )

    def test_visible_content_dropped_by_reader_is_rejected(self) -> None:
        fragments = (
            '<w:r><w:sym w:font="Wingdings" w:char="F0E0"/></w:r>',
            '<w:r><w:pgNum/></w:r>',
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            '<m:r><m:t>X</m:t></m:r></m:oMath>',
        )
        for fragment in fragments:
            with self.subTest(fragment=fragment[:25]):
                self.assert_rejected(lambda xml, f=fragment: xml.replace("</w:p>", f + "</w:p>", 1))

    def test_explicitly_visible_text_and_edit_permissions_are_valid(self) -> None:
        project_path = self.project()
        self.mutate_document(
            project_path,
            lambda xml: xml.replace(
                "</w:p>",
                '<w:r><w:rPr><w:vanish w:val="0"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r></w:p>'
                '<w:permStart w:id="1" w:edGrp="everyone"/><w:permEnd w:id="1"/>',
                1,
            ),
        )
        result = validate_project(project_path)
        self.assertTrue(result.valid, result.diagnostics.to_dict())

    def test_table_in_word_footer_is_rejected(self) -> None:
        project_path = self.project()
        template = project_path.parent / "template.docx"
        with zipfile.ZipFile(template) as archive:
            footer = archive.read("word/footer1.xml").decode("utf-8")
        table = "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>OCULTO</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        mutated = footer.replace("</w:ftr>", table + "</w:ftr>")
        hostile = project_path.parent / "hostile.docx"
        rewrite_zip(template, hostile, replacements={"word/footer1.xml": mutated.encode("utf-8")})
        os.replace(hostile, template)
        self.assertIn("DOCX-STRUCT-012", result_codes(validate_project(project_path)))

    def test_missing_page_size_is_a_diagnostic(self) -> None:
        import re as _re

        project_path = self.project()
        self.mutate_document(project_path, lambda xml: _re.sub(r"<w:pgSz[^>]*/>", "", xml, count=1))
        result = validate_project(project_path)
        self.assertIn("DOCX-STYLE-010", result_codes(result))

    def test_landscape_page_without_orient_attribute_is_landscape(self) -> None:
        project_path = self.project(orientation="LANDSCAPE")
        self.mutate_document(project_path, lambda xml: xml.replace(' w:orient="landscape"', "", 1))
        result = validate_project(project_path)
        self.assertTrue(result.valid, result.diagnostics.to_dict())
        self.assertEqual(result.template.template_orientation, "LANDSCAPE")

    def test_multibyte_heading_is_limited_in_bytes(self) -> None:
        columns = (("é" * 200, "VDNI"), ("Nombre", "VNOM"), ("Departamento", "DEPARTAMENTO"))
        result = validate_project(self.project(columns=columns))
        self.assertIn("DOCX-TABLE-009", result_codes(result))


class PlaceholderAndSqlRegressionTests(unittest.TestCase):
    def test_non_ascii_space_inside_placeholder_is_rejected(self) -> None:
        for text in ("{{ APP_USER}}", "{{\tAPP_USER}}", "{{FIELD: X}}"):
            with self.subTest(text=text):
                diagnostics = Diagnostics()
                validate_text_placeholders(
                    text,
                    allowed_kinds={"APP_USER", "FIELD"},
                    diagnostics=diagnostics,
                    location="Footer",
                )
                self.assertIn("TOKEN-003", codes(diagnostics))

    def test_ascii_spaces_inside_placeholder_remain_valid(self) -> None:
        diagnostics = Diagnostics()
        validate_text_placeholders(
            "{{ APP_USER }} {{FIELD : X}}",
            allowed_kinds={"APP_USER", "FIELD"},
            diagnostics=diagnostics,
            location="Footer",
        )
        self.assertFalse(diagnostics.has_errors, diagnostics.to_dict())

    def test_leading_comment_followed_by_ascii_whitespace_is_valid(self) -> None:
        for sql in ("/* cabecera */\nSELECT 1 x FROM dual", "-- x\n\n\tSELECT 1 x FROM dual"):
            with self.subTest(sql=sql):
                diagnostics = Diagnostics()
                validate_sql(sql, diagnostics)
                self.assertFalse(diagnostics.has_errors, diagnostics.to_dict())

    def test_non_ascii_whitespace_before_select_is_rejected(self) -> None:
        diagnostics = Diagnostics()
        validate_sql("/*c*/ SELECT 1 FROM dual", diagnostics)
        self.assertIn("SQL-006", codes(diagnostics))

    def test_apex_substitution_syntax_is_rejected_outside_literals(self) -> None:
        diagnostics = Diagnostics()
        validate_sql("SELECT 1 x FROM t WHERE c = &P1_X.", diagnostics)
        self.assertIn("SQL-013", codes(diagnostics))

    def test_substitution_text_inside_literal_or_comment_only_warns(self) -> None:
        for sql in ("SELECT '&P42_DNI.' x FROM dual", "SELECT 1 x FROM dual -- &APP_USER.", "SELECT 'AT&T.' x FROM dual"):
            with self.subTest(sql=sql):
                diagnostics = Diagnostics()
                validate_sql(sql, diagnostics)
                self.assertFalse(diagnostics.has_errors, diagnostics.to_dict())
                self.assertIn("SQL-015", codes(diagnostics))

    def test_into_and_update_inside_identifiers_or_literals_are_valid(self) -> None:
        for sql in (
            "SELECT into$flag AS x FROM t",
            "SELECT update#tipo AS x FROM t",
            'SELECT "INTO" AS x FROM t',
            "SELECT 'FOR UPDATE' AS x FROM t",
        ):
            with self.subTest(sql=sql):
                diagnostics = Diagnostics()
                validate_sql(sql, diagnostics)
                self.assertNotIn("SQL-014", codes(diagnostics))

    def test_reserved_name_as_field_points_to_builtin(self) -> None:
        diagnostics = Diagnostics()
        validate_text_placeholders(
            "{{FIELD:APP_USER}}", allowed_kinds={"FIELD"}, diagnostics=diagnostics, location="Footer"
        )
        self.assertIn("TOKEN-005", codes(diagnostics))

    def test_ampersand_in_literal_without_substitution_is_valid(self) -> None:
        diagnostics = Diagnostics()
        validate_sql("SELECT 'R&D' x FROM dual", diagnostics)
        self.assertFalse(diagnostics.has_errors, diagnostics.to_dict())

    def test_into_and_for_update_are_rejected(self) -> None:
        for sql in ("SELECT 1 INTO :X FROM dual", "SELECT * FROM t FOR UPDATE"):
            with self.subTest(sql=sql):
                diagnostics = Diagnostics()
                validate_sql(sql, diagnostics)
                self.assertIn("SQL-014", codes(diagnostics))

    def test_emitter_never_emits_a_raw_ampersand(self) -> None:
        expression = oracle_expression("R&D &P1_X. &", clob=True)
        self.assertNotIn("&", expression)
        self.assertEqual(expression.count("chr(38)"), 3)
        self.assertTrue(expression.startswith("to_clob("))
        self.assertEqual(oracle_expression("&", clob=False), "chr(38)")


class ProjectRegressionTests(TemporaryTestCase):
    def validate_with(self, **changes: object):
        project_path = self.project()
        self.rewrite_project(project_path, **changes)
        return validate_project(project_path)

    def test_style_keys_match_package_zones(self) -> None:
        for zone, key, value in (
            ("title", "background_color", "#FF0000"),
            ("footer", "background_color", "#00FF00"),
        ):
            with self.subTest(zone=zone, key=key):
                result = self.validate_with(style_overrides={zone: {key: value}})
                self.assertIn("PROJECT-083", result_codes(result))

    def test_table_body_alignment_is_ignored_with_warning(self) -> None:
        result = self.validate_with(style_overrides={"table_body": {"alignment": "END"}})
        self.assertTrue(result.valid, result.diagnostics.to_dict())
        self.assertIn("PROJECT-109", result_codes(result))
        self.assertNotIn("alignment", result.definition["styles"]["table_body"])

    def test_title_and_header_alias_together_are_rejected(self) -> None:
        result = self.validate_with(style_overrides={"title": {"font_size": 20}, "header": {"font_size": 10}})
        self.assertIn("PROJECT-107", result_codes(result))

    def test_title_is_limited_in_utf8_bytes(self) -> None:
        result = self.validate_with(title="ñ" * 200)
        self.assertIn("PROJECT-100", result_codes(result))

    def test_text_properties_reject_other_json_types(self) -> None:
        for key, value in (("title", None), ("title", {"x": 1}), ("file_name", 0), ("report_id", True)):
            with self.subTest(key=key, value=value):
                result = self.validate_with(**{key: value})
                self.assertFalse(result.valid)

    def test_duplicate_keys_and_deep_nesting_are_diagnostics(self) -> None:
        project_path = self.project()
        text = project_path.read_text(encoding="utf-8")
        project_path.write_text(text.replace('"title":', '"title": "A", "title":', 1), encoding="utf-8")
        self.assertIn("PROJECT-002", result_codes(validate_project(project_path)))

        project_path.write_text("[" * 100_000 + "]" * 100_000, encoding="utf-8")
        self.assertIn("PROJECT-002", result_codes(validate_project(project_path)))

    def test_reserved_apex_name_cannot_be_a_bind(self) -> None:
        project_path = self.project(query="select 1 vdni, 2 vnom, 3 departamento from dual where :APP_USER is not null")
        self.rewrite_project(
            project_path,
            bindings=[{"bind": "APP_USER", "item": "P42_USUARIO", "type": "VARCHAR2", "required": False}],
        )
        self.assertIn("PROJECT-017", result_codes(validate_project(project_path)))

    def test_alias_and_current_key_together_are_rejected(self) -> None:
        project_path = self.project()
        payload = json.loads(project_path.read_text(encoding="utf-8"))
        self.rewrite_project(project_path, binds=payload["bindings"])
        self.assertIn("PROJECT-105", result_codes(validate_project(project_path)))

    def test_weight_has_upper_bound(self) -> None:
        result = self.validate_with(
            column_widths=[{"column": "VNOM", "mode": "WEIGHT", "value": 1e300}]
        )
        self.assertIn("PROJECT-058", result_codes(result))

    def test_orientation_contradicting_docx_warns(self) -> None:
        result = self.validate_with(orientation="LANDSCAPE")
        self.assertTrue(result.valid)
        self.assertIn("PROJECT-102", result_codes(result))


class CompilerRegressionTests(TemporaryTestCase):
    def test_output_cannot_overwrite_project_inputs(self) -> None:
        project_path = self.project()
        source_sql = project_path.parent / "report.sql"
        renamed = project_path.parent / "apex_process.sql"
        os.replace(source_sql, renamed)
        self.rewrite_project(project_path, query_file="apex_process.sql")
        original = renamed.read_bytes()

        result = compile_project(project_path, project_path.parent)

        self.assertIn("IO-003", result_codes(result))
        self.assertEqual(renamed.read_bytes(), original)

    def test_directory_with_artifact_name_is_not_deleted(self) -> None:
        project_path = self.project()
        output = self.root / "out"
        keep = output / "template.json" / "keep.txt"
        keep.parent.mkdir(parents=True)
        keep.write_text("x", encoding="utf-8")

        result = compile_project(project_path, output)

        self.assertIn("IO-001", result_codes(result))
        self.assertTrue(keep.is_file())

    def test_interrupted_publication_restores_previous_output(self) -> None:
        project_path = self.project()
        output = self.root / "out"
        first = compile_project(project_path, output)
        self.assertTrue(first.valid)
        before = {path.name: path.read_bytes() for path in output.iterdir()}
        self.rewrite_project(project_path, title="Otro título")

        real_replace = os.replace
        calls = {"count": 0}

        def interrupted(source, destination):
            calls["count"] += 1
            if calls["count"] == 4:
                raise KeyboardInterrupt
            return real_replace(source, destination)

        with mock.patch.object(compiler_module.os, "replace", side_effect=interrupted):
            with self.assertRaises(KeyboardInterrupt):
                compile_project(project_path, output)

        after = {path.name: path.read_bytes() for path in output.iterdir()}
        self.assertEqual(after, before)


class CliRegressionTests(TemporaryTestCase):
    def test_json_output_survives_non_utf8_pipe(self) -> None:
        project_path = make_valid_project(self.root / "报告 Ω")
        environment = dict(os.environ, PYTHONIOENCODING="cp1252", PYTHONPATH=str(ROOT / "src"))
        completed = subprocess.run(
            [
                sys.executable,
                "-m",
                "corporate_report_compiler",
                "compile",
                "--project",
                str(project_path),
                "--output",
                str(self.root / "salida Ω"),
                "--json",
            ],
            capture_output=True,
            env=environment,
            check=False,
        )
        self.assertEqual(completed.returncode, 0, completed.stderr.decode("utf-8", "replace"))
        self.assertNotIn(b"Traceback", completed.stderr)
        payload = json.loads(completed.stdout.decode("utf-8"))
        self.assertTrue(payload["valid"])


if __name__ == "__main__":
    unittest.main()
