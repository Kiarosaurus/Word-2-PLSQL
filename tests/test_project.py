from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from corporate_report_compiler.diagnostics import Diagnostics
from corporate_report_compiler.docx_reader import read_template
from corporate_report_compiler.project import load_project, locate_project_inputs, read_sql

from tests.support import make_valid_project


class ProjectValidationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def load(self, project_path: Path):
        diagnostics = Diagnostics(strict=True)
        _raw, template_path, query_path = locate_project_inputs(project_path, diagnostics)
        if template_path is None or query_path is None:
            return diagnostics, None
        template = read_template(template_path, diagnostics)
        sql, binds = read_sql(query_path, diagnostics)
        model = load_project(project_path, template, sql, binds, diagnostics) if template is not None else None
        return diagnostics, model

    def rewrite(self, project_path: Path, **changes: object) -> None:
        payload = json.loads(project_path.read_text(encoding="utf-8"))
        payload.update(changes)
        project_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    def test_valid_project_loads_and_normalizes_contract(self) -> None:
        project_path = make_valid_project(self.root)
        diagnostics, model = self.load(project_path)
        self.assertFalse(diagnostics.has_errors, diagnostics.to_dict())
        self.assertIsNotNone(model)
        assert model is not None
        self.assertEqual(model.report_id, "PERSONAS")
        self.assertEqual(model.bindings[0]["bind"], "P_DNI")
        # APP_USER y GENERATED_AT son placeholders integrados que el package
        # resuelve directamente. Solo FIELD requiere un mapeo serializado.
        self.assertEqual(model.fields, ({
            "name": "OFFICE_NAME",
            "source": "ITEM",
            "item": "P42_OFFICE_NAME",
            "type": "VARCHAR2",
        },))
        self.assertEqual(model.column_widths[1], {"column": "VNOM", "mode": "AUTO"})

    def test_missing_and_surplus_bind_mappings_are_rejected(self) -> None:
        for bindings in (
            [],
            [
                {"bind": "P_DNI", "item": "P42_DNI", "type": "VARCHAR2"},
                {"bind": "EXTRA", "item": "P42_EXTRA", "type": "VARCHAR2"},
            ],
        ):
            with self.subTest(bindings=bindings):
                case = self.root / str(len(list(self.root.iterdir())))
                project_path = make_valid_project(case, project_overrides={"bindings": bindings})
                diagnostics, model = self.load(project_path)
                self.assertIsNone(model)
                self.assertTrue({"SQL-020", "SQL-021"} & {item.code for item in diagnostics.items})

    def test_required_date_bind_needs_format_mask(self) -> None:
        project_path = make_valid_project(
            self.root,
            project_overrides={
                "bindings": [{"bind": "P_DNI", "item": "P42_DATE", "type": "DATE"}],
            },
        )
        diagnostics, model = self.load(project_path)
        self.assertIsNone(model)
        self.assertIn("PROJECT-015", {item.code for item in diagnostics.items})

    def test_field_mapping_must_match_template_exactly(self) -> None:
        project_path = make_valid_project(self.root, project_overrides={"fields": []})
        diagnostics, model = self.load(project_path)
        self.assertIsNone(model)
        self.assertIn("TOKEN-020", {item.code for item in diagnostics.items})

    def test_all_columns_cannot_be_excluded(self) -> None:
        project_path = make_valid_project(
            self.root,
            project_overrides={"excluded_columns": ["VDNI", "VNOM", "DEPARTAMENTO"]},
        )
        diagnostics, model = self.load(project_path)
        self.assertIsNone(model)
        self.assertIn("PROJECT-043", {item.code for item in diagnostics.items})

    def test_invalid_width_contracts_are_rejected(self) -> None:
        cases = (
            ([{"column": "VNOM", "mode": "AUTO"}, {"column": "VDNI", "mode": "AUTO"}], "PROJECT-059"),
            ([{"column": "VDNI", "mode": "FIXED_PERCENT", "value": 99}], "PROJECT-060"),
            ([{"column": "NO_EXISTE", "mode": "WEIGHT", "value": 1}], "PROJECT-052"),
            ([{"column": "VNOM", "mode": "WEIGHT", "value": 0}], "PROJECT-058"),
        )
        for index, (widths, code) in enumerate(cases):
            with self.subTest(code=code):
                project_path = make_valid_project(
                    self.root / str(index),
                    project_overrides={"column_widths": widths},
                )
                diagnostics, model = self.load(project_path)
                self.assertIsNone(model)
                self.assertIn(code, {item.code for item in diagnostics.items})

    def test_page_size_and_auto_reserve_are_not_project_options(self) -> None:
        cases = (("paper_size", "A4", "PROJECT-076"), ("auto_reserve_pct", 25, "PROJECT-077"))
        for index, (key, value, code) in enumerate(cases):
            with self.subTest(key=key):
                project_path = make_valid_project(
                    self.root / f"obsolete-{index}",
                    project_overrides={key: value},
                )
                diagnostics, model = self.load(project_path)
                self.assertIsNone(model)
                self.assertIn(code, {item.code for item in diagnostics.items})

    def test_unknown_top_level_property_is_rejected(self) -> None:
        project_path = make_valid_project(self.root, project_overrides={"titlle": "Error tipográfico"})
        diagnostics, model = self.load(project_path)
        self.assertIsNone(model)
        self.assertIn("PROJECT-092", {item.code for item in diagnostics.items})

    def test_required_must_be_a_real_json_boolean(self) -> None:
        project_path = make_valid_project(
            self.root,
            project_overrides={
                "bindings": [
                    {"bind": "P_DNI", "item": "P42_DNI", "type": "VARCHAR2", "required": "false"},
                ],
            },
        )
        diagnostics, model = self.load(project_path)
        self.assertIsNone(model)
        self.assertIn("PROJECT-016", {item.code for item in diagnostics.items})

    def test_style_override_font_limits_match_runtime_zones(self) -> None:
        cases = (
            ("title", 7.5),
            ("table_header", 16.5),
            ("table_body", 14.5),
            ("footer", 12.5),
        )
        for index, (zone, size) in enumerate(cases):
            with self.subTest(zone=zone):
                project_path = make_valid_project(
                    self.root / f"style-{index}",
                    project_overrides={"style_overrides": {zone: {"font_size": size}}},
                )
                diagnostics, model = self.load(project_path)
                self.assertIsNone(model)
                self.assertIn("PROJECT-090", {item.code for item in diagnostics.items})

    def test_title_override_must_remain_centered(self) -> None:
        project_path = make_valid_project(
            self.root,
            project_overrides={"style_overrides": {"title": {"alignment": "START"}}},
        )
        diagnostics, model = self.load(project_path)
        self.assertIsNone(model)
        self.assertIn("PROJECT-093", {item.code for item in diagnostics.items})

    def test_page_item_pattern_matches_runtime(self) -> None:
        project_path = make_valid_project(
            self.root,
            project_overrides={
                "bindings": [{"bind": "P_DNI", "item": "P42_DNI$", "type": "VARCHAR2"}],
            },
        )
        diagnostics, model = self.load(project_path)
        self.assertIsNone(model)
        self.assertIn("PROJECT-013", {item.code for item in diagnostics.items})

    def test_page_item_runtime_length_limit_is_enforced_locally(self) -> None:
        project_path = make_valid_project(
            self.root,
            project_overrides={
                "bindings": [
                    {
                        "bind": "P_DNI",
                        "item": "P" + ("9" * 126) + "_DNI",
                        "type": "VARCHAR2",
                    }
                ],
            },
        )
        diagnostics, model = self.load(project_path)
        self.assertIsNone(model)
        self.assertIn("PROJECT-013", {item.code for item in diagnostics.items})

    def test_excluded_widths_do_not_consume_visible_budget(self) -> None:
        project_path = make_valid_project(
            self.root,
            project_overrides={
                "excluded_columns": ["VDNI"],
                "column_widths": [
                    {"column": "VDNI", "mode": "AUTO"},
                    {"column": "VNOM", "mode": "AUTO"},
                    {"column": "DEPARTAMENTO", "mode": "FIXED_PERCENT", "value": 20},
                ],
            },
        )
        diagnostics, model = self.load(project_path)
        self.assertFalse(diagnostics.has_errors, diagnostics.to_dict())
        self.assertIsNotNone(model)

    def test_unknown_nested_property_is_rejected(self) -> None:
        project_path = make_valid_project(
            self.root,
            project_overrides={
                "bindings": [
                    {"bind": "P_DNI", "item": "P42_DNI", "type": "VARCHAR2", "requred": True},
                ],
            },
        )
        diagnostics, model = self.load(project_path)
        self.assertIsNone(model)
        self.assertIn("PROJECT-094", {item.code for item in diagnostics.items})

    def test_non_finite_json_numbers_are_rejected(self) -> None:
        project_path = make_valid_project(self.root)
        payload = project_path.read_text(encoding="utf-8").replace('"value": 14', '"value": NaN')
        project_path.write_text(payload, encoding="utf-8")
        diagnostics, model = self.load(project_path)
        self.assertIsNone(model)
        self.assertIn("PROJECT-002", {item.code for item in diagnostics.items})

        project_path = make_valid_project(self.root / "overflow")
        payload = project_path.read_text(encoding="utf-8").replace('"value": 14', '"value": 1e10000')
        project_path.write_text(payload, encoding="utf-8")
        diagnostics, model = self.load(project_path)
        self.assertIsNone(model)
        self.assertIn("PROJECT-056", {item.code for item in diagnostics.items})

    def test_extreme_json_integers_return_diagnostics_instead_of_crashing(self) -> None:
        huge = 10 ** 4000
        cases = (
            (
                {"fields": [{"name": "OFFICE_NAME", "source": "CONSTANT", "value": huge}]},
                "PROJECT-039",
            ),
            ({"style_overrides": {"title": {"font_size": huge}}}, "PROJECT-088"),
            (
                {"column_widths": [{"column": "VDNI", "mode": "WEIGHT", "value": huge}]},
                "PROJECT-056",
            ),
        )
        for index, (overrides, code) in enumerate(cases):
            with self.subTest(code=code):
                project_path = make_valid_project(
                    self.root / f"huge-number-{index}",
                    project_overrides=overrides,
                )
                diagnostics, model = self.load(project_path)
                self.assertIsNone(model)
                self.assertIn(code, {item.code for item in diagnostics.items})

    def test_project_cannot_reference_files_outside_its_directory(self) -> None:
        project_path = make_valid_project(self.root / "project")
        external = self.root / "external.sql"
        external.write_text("select 1 from dual\n", encoding="utf-8")
        self.rewrite(project_path, query_file="../external.sql")
        diagnostics, model = self.load(project_path)
        self.assertIsNone(model)
        self.assertIn("PROJECT-005", {item.code for item in diagnostics.items})

    def test_runtime_text_length_limits_are_enforced_locally(self) -> None:
        cases = (
            ({"title": "T" * 256}, "PROJECT-100"),
            ({"file_name": "f" * 181}, "PROJECT-101"),
            (
                {"bindings": [{"bind": "P_DNI", "item": "P42_DNI", "type": "VARCHAR2", "format_mask": "X" * 4001}]},
                "PROJECT-098",
            ),
            (
                {"fields": [{"name": "OFFICE_NAME", "source": "CONSTANT", "value": "ñ" * 2001}]},
                "PROJECT-099",
            ),
        )
        for index, (overrides, code) in enumerate(cases):
            with self.subTest(code=code):
                project_path = make_valid_project(
                    self.root / f"text-limit-{index}",
                    project_overrides=overrides,
                )
                diagnostics, model = self.load(project_path)
                self.assertIsNone(model)
                self.assertIn(code, {item.code for item in diagnostics.items})


if __name__ == "__main__":
    unittest.main()
