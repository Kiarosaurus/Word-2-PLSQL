from __future__ import annotations

from contextlib import redirect_stderr, redirect_stdout
import io
import json
from pathlib import Path
import tempfile
import unittest

from corporate_report_compiler import compile_project, validate_project
from corporate_report_compiler.cli import main

from tests.support import make_valid_project


class CompilerIntegrationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        self.project = make_valid_project(self.root / "project")

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def test_validate_is_read_only_and_compile_emits_exact_contract(self) -> None:
        output = self.root / "compiled"
        validated = validate_project(self.project)
        self.assertTrue(validated.valid, validated.diagnostics.to_dict())
        self.assertFalse(output.exists())

        result = compile_project(self.project, output)
        self.assertTrue(result.valid, result.diagnostics.to_dict())
        self.assertEqual(
            {path.name for path in result.artifacts},
            {"template.json", "validation.json", "apex_process.sql"},
        )
        self.assertEqual(
            json.loads((output / "template.json").read_text(encoding="utf-8")),
            result.definition,
        )
        validation = json.loads((output / "validation.json").read_text(encoding="utf-8"))
        self.assertTrue(validation["valid"])

    def test_invalid_project_does_not_create_or_overwrite_output(self) -> None:
        output = self.root / "compiled"
        output.mkdir()
        sentinel = output / "template.json"
        sentinel.write_text("last-known-good", encoding="utf-8")
        self.project.with_name("report.sql").write_text("delete from persona", encoding="utf-8")

        result = compile_project(self.project, output)
        self.assertFalse(result.valid)
        self.assertEqual(result.artifacts, ())
        self.assertEqual(sentinel.read_text(encoding="utf-8"), "last-known-good")
        self.assertFalse((output / "apex_process.sql").exists())

    def test_generated_process_matches_download_query_public_signature(self) -> None:
        result = compile_project(self.project, self.root / "compiled")
        self.assertTrue(result.valid, result.diagnostics.to_dict())
        process = (self.root / "compiled" / "apex_process.sql").read_text(encoding="utf-8").lower()
        self.assertIn("pkg_corporate_reports.download_query", process)
        self.assertNotIn("p_paper_size", process)
        self.assertNotIn("p_auto_reserve_pct", process)

    def test_cli_validate_json_and_compile(self) -> None:
        stdout = io.StringIO()
        stderr = io.StringIO()
        with redirect_stdout(stdout), redirect_stderr(stderr):
            code = main(["validate", "--project", str(self.project), "--json"])
        self.assertEqual(code, 0, stderr.getvalue())
        self.assertTrue(json.loads(stdout.getvalue())["valid"])

        output = self.root / "cli-output"
        with redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
            code = main(
                ["compile", "--project", str(self.project), "--output", str(output)]
            )
        self.assertEqual(code, 0)
        self.assertTrue((output / "template.json").is_file())


if __name__ == "__main__":
    unittest.main()
