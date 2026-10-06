"""Organización proyectos/<nombre>/ con la carpeta generado/."""

from __future__ import annotations

import json
import shutil
import tempfile
import unittest
from pathlib import Path

from corporate_report_compiler.cli import main
from corporate_report_compiler.compiler import compile_project, validate_project
from corporate_report_compiler.gui import default_output
from corporate_report_compiler.workspace import files_to_replace, import_docx, workspace_outputs


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates" / "entidades_portrait.docx"
EXAMPLE = ROOT / "proyectos" / "demo" / "entidades" / "generado" / "entidades.report.json"


class ImportDocxTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.base = Path(self.temporary.name)
        self.projects = self.base / "proyectos"
        self.source = self.base / "otra_carpeta" / "ventas.docx"
        self.source.parent.mkdir()
        shutil.copyfile(TEMPLATE, self.source)

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def test_docx_is_copied_into_folder_with_its_name(self) -> None:
        result = import_docx(self.source, projects_root=self.projects, page="42")

        self.assertFalse(result.diagnostics.has_errors, result.diagnostics.to_dict())
        folder = self.projects / "ventas"
        self.assertEqual((folder / "ventas.docx").read_bytes(), self.source.read_bytes())
        self.assertTrue((folder / "generado" / "ventas.sql").is_file())
        project = json.loads((folder / "generado" / "ventas.report.json").read_text(encoding="utf-8"))
        self.assertEqual(project["template"], "../ventas.docx")
        self.assertEqual(project["query_file"], "ventas.sql")
        self.assertIn({"name": "UNIDAD", "source": "ITEM", "item": "P42_UNIDAD", "type": "VARCHAR2"}, project["fields"])
        self.assertTrue(validate_project(folder / "generado" / "ventas.report.json").valid)

    def test_compile_splits_upload_file_from_generated_files(self) -> None:
        import_docx(self.source, projects_root=self.projects, page="42")
        project = self.projects / "ventas" / "generado" / "ventas.report.json"
        output, process_directory = workspace_outputs(project)

        result = compile_project(project, output, process_directory=process_directory)

        self.assertTrue(result.valid, result.diagnostics.to_dict())
        folder = self.projects / "ventas"
        self.assertEqual(
            {path.relative_to(folder).as_posix() for path in result.artifacts},
            {"apex_process.sql", "generado/template.json", "generado/validation.json"},
        )

    def test_reupload_requires_replace_and_keeps_backups_and_other_files(self) -> None:
        import_docx(self.source, projects_root=self.projects, page="42")
        folder = self.projects / "ventas"
        sql = folder / "generado" / "ventas.sql"
        sql.write_text("select edited from dual\n", encoding="utf-8")
        notes = folder / "notas.txt"
        notes.write_text("no tocar", encoding="utf-8")
        extra = folder / "generado" / "otro.txt"
        extra.write_text("tampoco", encoding="utf-8")

        refused = import_docx(self.source, projects_root=self.projects, page="42")
        self.assertIn("GUIDE-002", {item.code for item in refused.diagnostics.items})
        self.assertEqual(sql.read_text(encoding="utf-8"), "select edited from dual\n")

        result = import_docx(self.source, projects_root=self.projects, page="42", replace=True)

        self.assertFalse(result.diagnostics.has_errors, result.diagnostics.to_dict())
        self.assertEqual(
            (folder / "generado" / "ventas.sql.bak").read_text(encoding="utf-8"),
            "select edited from dual\n",
        )
        self.assertNotEqual(sql.read_text(encoding="utf-8"), "select edited from dual\n")
        self.assertTrue((folder / "generado" / "ventas.report.json.bak").is_file())
        self.assertTrue((folder / "generado" / "ventas.docx.bak").is_file())
        self.assertEqual(notes.read_text(encoding="utf-8"), "no tocar")
        self.assertEqual(extra.read_text(encoding="utf-8"), "tampoco")

    def test_reimporting_the_copy_in_place_does_not_back_up_the_docx(self) -> None:
        import_docx(self.source, projects_root=self.projects)
        in_place = self.projects / "ventas" / "ventas.docx"

        self.assertNotIn(in_place, files_to_replace(in_place, self.projects))
        result = import_docx(in_place, projects_root=self.projects, replace=True)

        self.assertFalse(result.diagnostics.has_errors, result.diagnostics.to_dict())
        self.assertFalse((self.projects / "ventas" / "generado" / "ventas.docx.bak").exists())
        self.assertTrue(in_place.is_file())

    def test_invalid_docx_changes_nothing(self) -> None:
        broken = self.base / "roto.docx"
        broken.write_bytes(b"no es un zip")

        result = import_docx(broken, projects_root=self.projects)

        self.assertTrue(result.diagnostics.has_errors)
        self.assertFalse((self.projects / "roto").exists())

    def test_cli_new_and_compile_without_output(self) -> None:
        self.assertEqual(main(["new", "--docx", str(self.source), "--projects-dir", str(self.projects), "--page", "42"]), 0)
        self.assertEqual(main(["new", "--docx", str(self.source), "--projects-dir", str(self.projects)]), 1)
        project = self.projects / "ventas" / "generado" / "ventas.report.json"
        self.assertEqual(main(["compile", "--project", str(project)]), 0)
        self.assertTrue((self.projects / "ventas" / "apex_process.sql").is_file())


class LayoutTests(unittest.TestCase):
    def test_example_project_uses_the_layout(self) -> None:
        folder = EXAMPLE.parent.parent
        for name in ("entidades.docx", "apex_process.sql", "datos_prueba.sql"):
            self.assertTrue((folder / name).is_file(), name)
        # La guía común de las demos vive en proyectos/demo/, junto a ambos proyectos.
        self.assertTrue((folder.parent / "README.docx").is_file())
        self.assertTrue((folder.parent / "estado_cuenta" / "estado_cuenta.docx").is_file())
        for name in ("entidades.sql", "entidades.report.json", "template.json", "validation.json"):
            self.assertTrue((EXAMPLE.parent / name).is_file(), name)

    def test_example_outputs_match_a_fresh_compilation(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            copy = Path(directory) / "entidades"
            shutil.copytree(EXAMPLE.parent.parent, copy)
            project = copy / "generado" / "entidades.report.json"
            output, process_directory = workspace_outputs(project)
            result = compile_project(project, output, process_directory=process_directory)
            self.assertTrue(result.valid, result.diagnostics.to_dict())
            for relative in ("apex_process.sql", "generado/template.json", "generado/validation.json"):
                self.assertEqual(
                    (copy / relative).read_bytes(),
                    (EXAMPLE.parent.parent / relative).read_bytes(),
                    relative,
                )

    def test_default_output_inside_generated_folder(self) -> None:
        self.assertEqual(default_output(EXAMPLE), EXAMPLE.parent)

    def test_paths_cannot_escape_the_project_folder(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            copy = Path(directory) / "proyectos" / "entidades"
            shutil.copytree(EXAMPLE.parent.parent, copy)
            shutil.copyfile(copy / "entidades.docx", Path(directory) / "proyectos" / "fuera.docx")
            project = copy / "generado" / "entidades.report.json"
            raw = json.loads(project.read_text(encoding="utf-8"))
            raw["template"] = "../../fuera.docx"
            project.write_text(json.dumps(raw), encoding="utf-8")
            result = validate_project(project)
            self.assertIn("PROJECT-005", {item.code for item in result.diagnostics.items})


if __name__ == "__main__":
    unittest.main()
