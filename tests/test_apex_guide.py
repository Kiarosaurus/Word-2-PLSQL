from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

from corporate_report_compiler.apex_guide import build_apex_guide, page_item_rows, write_skeleton
from corporate_report_compiler.compiler import compile_project, validate_project
from corporate_report_compiler.gui import _format_result, default_output


ROOT = Path(__file__).resolve().parents[1]
EXAMPLE = ROOT / "proyectos" / "entidades" / "generado" / "entidades.report.json"


class ApexGuideTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def test_guide_lists_files_items_and_locations(self) -> None:
        result = compile_project(EXAMPLE, self.root / "out")
        self.assertTrue(result.valid)
        guide = build_apex_guide(result)
        for expected in (
            "sql/pkg_corporate_reports.sql",
            "SQL Workshop > SQL Scripts",
            str(self.root / "out" / "apex_process.sql"),
            "Execute Code",
            "When Button Pressed = DOWNLOAD_ENTIDADES",
            "P42_DNI",
            "P42_DEPARTAMENTO_ID",
            "máscara DD/MM/YYYY",
            "P0_REPORT_FORMAT",
            "PDF, XLSX",
            "AUTO, PORTRAIT, LANDSCAPE",
            "VDIREC_ACTUAL",
            "NO se sube",
            "Reload on Submit: Always",
        ):
            self.assertIn(expected, guide)
        self.assertIn("0 (Global Page)", {row[1] for row in page_item_rows(result)})

    def test_guide_is_empty_for_invalid_result(self) -> None:
        broken = self.root / "x.report.json"
        broken.write_text("{}", encoding="utf-8")
        self.assertEqual(build_apex_guide(validate_project(broken)), "")

    def test_gui_text_shows_guide_only_after_compiling(self) -> None:
        validated = _format_result(validate_project(EXAMPLE))
        self.assertNotIn("QUÉ SUBIR A APEX", validated)
        compiled = _format_result(compile_project(EXAMPLE, self.root / "out"), compiled=True)
        self.assertIn("QUÉ SUBIR A APEX", compiled)

    def test_default_output_drops_report_suffix(self) -> None:
        self.assertEqual(default_output(Path("C:/p/ventas.report.json")), Path("C:/p/build/ventas"))

    def test_skeleton_maps_docx_markers_to_page_items(self) -> None:
        docx = self.root / "entidades.docx"
        shutil.copyfile(ROOT / "proyectos" / "entidades" / "entidades.docx", docx)

        diagnostics, skeleton, files = write_skeleton(docx, page="42")

        self.assertFalse(diagnostics.has_errors, diagnostics.to_dict())
        self.assertEqual({path.name for path in files}, {"entidades.report.json", "entidades.sql"})
        items = {binding["item"] for binding in skeleton.project["bindings"]}
        self.assertIn("F_VDNI", {binding["bind"] for binding in skeleton.project["bindings"]})
        self.assertIn("P42_VDNI", items)
        self.assertIn("P42_VNRO_TLF1", items)
        result = validate_project(files[0])
        self.assertTrue(result.valid, result.diagnostics.to_dict())

    def test_skeleton_with_placeholder_page_and_no_overwrite(self) -> None:
        docx = self.root / "entidades.docx"
        shutil.copyfile(ROOT / "proyectos" / "entidades" / "entidades.docx", docx)
        _diagnostics, skeleton, files = write_skeleton(docx)
        self.assertTrue(all(binding["item"].startswith("PXX_") for binding in skeleton.project["bindings"]))
        before = files[0].read_bytes()

        diagnostics, _skeleton, again = write_skeleton(docx)

        self.assertEqual(again, ())
        self.assertIn("GUIDE-001", {item.code for item in diagnostics.items})
        self.assertEqual(files[0].read_bytes(), before)


class SkeletonEdgeCaseTests(unittest.TestCase):
    def test_reserved_and_colliding_aliases_still_validate(self) -> None:
        from tests.support import build_valid_docx

        columns = (
            ("Fecha", "DATE"), ("Nivel", "LEVEL"), ("Destino", "INTO"), ("Usuario", "APP_USER"),
            ("Monto 1", "MONTO$"), ("Monto 2", "MONTO#"), ("Monto 3", "MONTO_"),
        )
        with tempfile.TemporaryDirectory() as directory:
            docx = build_valid_docx(Path(directory) / "Reporte_SUNAT.docx", columns=columns, include_field=False)
            diagnostics, skeleton, files = write_skeleton(docx, page="42")
            self.assertFalse(diagnostics.has_errors, diagnostics.to_dict())
            items = [binding["item"] for binding in skeleton.project["bindings"]]
            self.assertEqual(len(items), len(set(items)))
            self.assertEqual(skeleton.project["title"], "Reporte SUNAT")
            result = validate_project(files[0])
            self.assertTrue(result.valid, result.diagnostics.to_dict())


class GuiOutputTests(unittest.TestCase):
    def test_output_follows_project_unless_chosen_for_it(self) -> None:
        try:
            import tkinter as tk

            root = tk.Tk()
        except Exception as exc:  # pragma: no cover - entorno sin Tk
            self.skipTest(f"Tk no disponible: {exc}")
        root.withdraw()
        try:
            from corporate_report_compiler.gui import CompilerApplication

            app = CompilerApplication(root)
            app._set_project(Path("C:/a/ventas.report.json"))
            self.assertEqual(Path(app.output.get()), Path("C:/a/build/ventas"))
            app._output_chosen_for = app.project.get()
            app.output.set("C:/elegida")
            app.project.set("C:/b/compras.report.json")
            self.assertFalse(app._output_is_custom())
            app._set_project(Path("C:/b/compras.report.json"))
            self.assertEqual(Path(app.output.get()), Path("C:/b/build/compras"))
        finally:
            root.destroy()


class PortableShortcutTests(unittest.TestCase):
    def test_shortcut_has_no_absolute_project_path_or_machine_data(self) -> None:
        import base64
        import os as _os
        import re as _re

        link = ROOT / "Compilador de reportes APEX.lnk"
        data = link.read_bytes()
        self.assertEqual(data[:4], b"L\0\0\0")
        text = data.decode("utf-16-le", "ignore") + data[1:].decode("utf-16-le", "ignore") + data.decode("latin-1")
        self.assertNotIn(str(ROOT), text)
        self.assertNotIn(b"\x03\x00\x00\xa0", data, "TrackerDataBlock con nombre del equipo")
        self.assertNotIn(b"\x09\x00\x00\xa0", data, "PropertyStoreDataBlock con el SID del autor")
        self.assertNotIn("S-1-5-21-", text)
        machine = _os.environ.get("COMPUTERNAME")
        if machine:
            self.assertNotIn(machine.lower(), text.lower())
        unicode_text = data.decode("utf-16-le", "ignore") + data[1:].decode("utf-16-le", "ignore")
        encoded = _re.search(r"-Enc ([A-Za-z0-9+/=]+)", unicode_text)
        self.assertIsNotNone(encoded)
        script = base64.b64decode(encoded.group(1)).decode("utf-16-le")
        self.assertIn("GetStartupInfoW", script)
        self.assertIn("Abrir compilador.bat", script)


if __name__ == "__main__":
    unittest.main()
