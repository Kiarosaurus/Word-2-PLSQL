"""Tablas de los FROM de un proyecto y su esquema / DB link."""

from __future__ import annotations

import json
import re
import shutil
import tempfile
import unittest
from pathlib import Path

from corporate_report_compiler.compiler import validate_project
from corporate_report_compiler.sources import (common_source, function_spans, project_functions, project_tables,
                                               qualified, restore_function_sources, set_function_source,
                                               set_table_source, table_spans)
from corporate_report_compiler.workspace import import_docx

from tests.test_reports_model import FIXTURE, build_docx


class TableSpansTests(unittest.TestCase):
    def test_finds_tables_of_from_and_join_only(self) -> None:
        sql = """select a.x, extract(year from a.f), trim(leading '0' from b.c),
                        (select max(z) from zz) m, 'from texto' t
                   from pgpago a, bdatos.presol@DL_11G b
                   left join cfperofi c on c.id = a.id
                   join table(f(1)) t on 1 = 1
                  where exists (select 1 from dual)  -- from comentario
                 union all select * from Otra.TABLA o"""
        self.assertEqual([(span.table, span.text) for span in table_spans(sql)], [
            ("ZZ", "zz"), ("PGPAGO", "pgpago"), ("PRESOL", "bdatos.presol@DL_11G"),
            ("CFPEROFI", "cfperofi"), ("TABLA", "Otra.TABLA"),
        ])

    def test_plsql_select_into(self) -> None:
        code = "function f return char is x char; begin select a into x from meses where n = 1; return x; end;"
        self.assertEqual([span.table for span in table_spans(code)], ["MESES"])

    def test_qualified(self) -> None:
        self.assertEqual(qualified("pgpago", "bdatos", "dl_11g"), "BDATOS.PGPAGO@DL_11G")
        self.assertEqual(qualified("pgpago", "", "dl_11g"), "PGPAGO@DL_11G")
        self.assertEqual(qualified("pgpago"), "PGPAGO")

    def test_external_function_spans_ignore_builtins_comments_and_literals(self) -> None:
        code = """function F return varchar2 is begin
          x := mes_anterior(:MES); y := Bdatos.fecha_titulo@DL_11G(1);
          z := nvl(to_char(1), 'fake(2)'); -- comentario(3)
          return helper(1); end;"""
        self.assertEqual([(span.function, span.text) for span in function_spans(code, {"F", "HELPER"})], [
            ("MES_ANTERIOR", "mes_anterior"), ("FECHA_TITULO", "Bdatos.fecha_titulo@DL_11G")])


class ProjectSourcesTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        docx = build_docx(self.root / "boleta.docx", ["{{COLUMN:NRO_CUOTA}}", "{{COLUMN:F_PEN}}", "{{COLUMN:F_TOT_CUOTA}}"],
                          header="{{FIELD:F_DOBLE}}")
        result = import_docx(docx, projects_root=self.root / "proyectos", page="71", reports=FIXTURE)
        self.assertFalse(result.diagnostics.has_errors, result.diagnostics.to_dict())
        self.project = result.paths.project
        self.generated = self.project.parent

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def test_tables_of_queries_get_schema_and_db_link(self) -> None:
        tables = {table.table: table for table in project_tables(self.project)}
        self.assertIn("PRUEBAP_ALUMNO_CUOTA", tables)
        self.assertEqual(common_source(self.project), ("", ""))

        change = set_table_source(self.project, ["PRUEBAP_ALUMNO_CUOTA"], "bdatos", "dl_11g")
        self.assertIsNone(change.error)
        self.assertEqual(change.tables, ["PRUEBAP_ALUMNO_CUOTA"])
        query = (self.generated / "q_q_cuotas.sql").read_text(encoding="utf-8")
        self.assertIn("from BDATOS.PRUEBAP_ALUMNO_CUOTA@DL_11G c", query)
        self.assertEqual(common_source(self.project), ("BDATOS", "DL_11G"))
        self.assertTrue(validate_project(self.project).valid)

        self.assertIn("No hay cambios", set_table_source(self.project, ["PRUEBAP_ALUMNO_CUOTA"],
                                                         "bdatos", "dl_11g").error)
        removed = set_table_source(self.project, ["PRUEBAP_ALUMNO_CUOTA"])
        self.assertIsNone(removed.error)
        self.assertIn("from PRUEBAP_ALUMNO_CUOTA c", (self.generated / "q_q_cuotas.sql").read_text(encoding="utf-8"))
        self.assertIn("no es un nombre de esquema", set_table_source(self.project, None, "mal esquema", "").error)

    def test_update_docx_keeps_adjustments_and_adds_what_is_new(self) -> None:
        from docx import Document

        from corporate_report_compiler.workspace import rename_items, update_docx

        set_table_source(self.project, ["PRUEBAP_ALUMNO_CUOTA"], "BDATOS", "DL_11G")
        rename_items(self.project, {item: item.replace("P71_", "P80_") for item in
                                    [entry["item"] for entry in json.loads(self.project.read_text(encoding="utf-8"))
                                     ["parameters"]]} or {})
        before = self.project.read_text(encoding="utf-8")
        word = self.project.parent.parent / "boleta.docx"

        same = update_docx(self.project, word)                      # mismo Word: nada más cambia
        self.assertIsNotNone(same.docx)
        self.assertEqual(self.project.read_text(encoding="utf-8"), before)
        self.assertEqual(same.added_queries, [])

        document = Document(str(word))
        document.add_paragraph("Vacía: {{FIELD:NRO_CUOTA2}}")       # dato de una consulta que no estaba
        changed = self.root / "boleta_v2.docx"
        document.save(str(changed))
        update = update_docx(self.project, changed)
        self.assertEqual(update.added_queries, ["q_q_vacia.sql"])
        self.assertEqual(update.qualified_tables, ["PRUEBAP_ALUMNO_CUOTA"])
        self.assertIn("BDATOS.PRUEBAP_ALUMNO_CUOTA@DL_11G", (self.generated / "q_q_vacia.sql").read_text(encoding="utf-8"))
        raw = json.loads(self.project.read_text(encoding="utf-8"))
        self.assertIn("Q_VACIA", raw["queries"])
        self.assertTrue(all(entry["item"].startswith("P80_") for entry in raw["parameters"]))   # se conservan
        self.assertTrue((self.project.parent.parent / "boleta.docx.bak").exists() or update.backup is not None)
        self.assertTrue(validate_project(self.project).valid)

        broken = Document(str(changed))
        broken.add_paragraph("{{FIELD:mal formado!}}")
        bad = self.root / "boleta_mal.docx"
        broken.save(str(bad))
        snapshot = self.project.read_text(encoding="utf-8")
        failed = update_docx(self.project, bad)                     # Word con errores: no se toca nada
        self.assertIsNone(failed.docx)
        self.assertTrue(failed.diagnostics.has_errors)
        self.assertEqual(self.project.read_text(encoding="utf-8"), snapshot)

    def test_formula_tables_are_written_to_codigo_reports(self) -> None:
        raw = json.loads(self.project.read_text(encoding="utf-8"))
        raw["reports_code"] = "codigo_reports"
        self.project.write_text(json.dumps(raw), encoding="utf-8")
        code_dir = self.generated / "codigo_reports"
        code_dir.mkdir(exist_ok=True)
        (code_dir / "cf_dobleformula.sql").write_text(
            "function CF_DOBLEFormula return Number is x number;\n"
            "begin\n  select count(*) into x from meses;\n  return x;\nend;\n", encoding="utf-8")
        tables = {table.table: table for table in project_tables(self.project)}
        self.assertEqual(tables["MESES"].places, ["fórmula cf_dobleformula"])

        change = set_table_source(self.project, ["MESES"], "", "DL_11G")
        self.assertIsNone(change.error)
        self.assertIn("from MESES@DL_11G;", (code_dir / "cf_dobleformula.sql").read_text(encoding="utf-8"))
        result = validate_project(self.project)
        self.assertTrue(result.valid, result.diagnostics.to_dict())
        body = next(unit["body"] for unit in result.definition["reports"]["units"] if unit["name"] == "cf_dobleformula")
        self.assertIn("MESES@DL_11G", body)

    def test_external_functions_are_written_to_codigo_reports(self) -> None:
        raw = json.loads(self.project.read_text(encoding="utf-8"))
        raw["reports_code"] = "codigo_reports"
        self.project.write_text(json.dumps(raw), encoding="utf-8")
        code_dir = self.generated / "codigo_reports"
        code_dir.mkdir(exist_ok=True)
        source = code_dir / "cf_dobleformula.sql"
        source.write_text(
            "function CF_DOBLEFormula return Number is\n"
            "begin\n  return mes_anterior(:NRO_CUOTA) + nvl(1, 0);\nend;\n", encoding="utf-8")
        found = {item.function: item for item in project_functions(self.project)}
        self.assertEqual(found["MES_ANTERIOR"].forms, {"mes_anterior": 1})
        change = set_function_source(self.project, ["MES_ANTERIOR"], "BDATOS", "DL_11G")
        self.assertIsNone(change.error)
        self.assertEqual(change.functions, ["MES_ANTERIOR"])
        self.assertIn("BDATOS.MES_ANTERIOR@DL_11G(:NRO_CUOTA)", source.read_text(encoding="utf-8"))
        self.assertEqual(project_functions(self.project)[0].forms, {"BDATOS.MES_ANTERIOR@DL_11G": 1})
        checked = validate_project(self.project)
        self.assertTrue(checked.valid)
        self.assertNotIn("DL_11G", " ".join(item.message for item in checked.diagnostics.items))
        self.assertIn("No hay cambios", set_function_source(self.project, ["MES_ANTERIOR"],
                                                              "BDATOS", "DL_11G").error)
        link_removed = set_function_source(self.project, ["MES_ANTERIOR"])
        self.assertIsNone(link_removed.error)
        self.assertIn("BDATOS.MES_ANTERIOR(:NRO_CUOTA)", source.read_text(encoding="utf-8"))
        removed = set_function_source(self.project, ["MES_ANTERIOR"], prefix_role="schema")
        self.assertIsNone(removed.error)
        self.assertIn("MES_ANTERIOR(:NRO_CUOTA)", source.read_text(encoding="utf-8"))
        self.assertIn("no es un nombre de esquema", set_function_source(self.project, None, "mal esquema").error)

    def test_query_function_source_reaches_compiler_and_restores_from_rdf(self) -> None:
        original = next((Path(__file__).resolve().parents[1] / "proyectos").glob("PSR308W_1I*"))
        copied = self.root / "PSR308W_1I"
        shutil.copytree(original, copied)
        project = copied / "generado" / "PSR308W_1I.report.json"
        query = copied / "generado" / "q_q_datos.sql"
        function = "SF_FSRYC_LIQUIDEZ_REM"
        before = re.sub(r"(?:[A-Za-z][A-Za-z0-9_$#]*\.)?SF_FSRYC_LIQUIDEZ_REM(?:@[A-Za-z0-9_$#.]+)?(?=\()",
                        function, query.read_text(encoding="utf-8"))
        query.write_text(before, encoding="utf-8")

        found = {item.function: item for item in project_functions(project)}
        self.assertEqual(found[function].places, ["q_q_datos.sql"])
        changed = set_function_source(project, [function], "BDATOS", "DL_11G")
        self.assertIsNone(changed.error)
        expected = "BDATOS.SF_FSRYC_LIQUIDEZ_REM@DL_11G(:cod) LIQ_DISP"
        self.assertIn(expected, query.read_text(encoding="utf-8"))
        checked = validate_project(project, strict=False)
        self.assertTrue(checked.valid, checked.diagnostics.to_dict())
        self.assertIn(expected, checked.definition["queries"]["Q_DATOS"]["sql"])

        restored = restore_function_sources(project, [function])
        self.assertIsNone(restored.error)
        self.assertEqual(query.read_text(encoding="utf-8"), before)
        self.assertEqual(len(restored.backups), 1)
        self.assertIn(expected, restored.backups[0].read_text(encoding="utf-8"))

    def test_qualified_function_can_get_link_and_explicit_package_schema(self) -> None:
        raw = json.loads(self.project.read_text(encoding="utf-8"))
        raw["reports_code"] = "codigo_reports"
        self.project.write_text(json.dumps(raw), encoding="utf-8")
        code_dir = self.generated / "codigo_reports"
        code_dir.mkdir(exist_ok=True)
        source = code_dir / "cf_dobleformula.sql"
        original = "function CF_DOBLEFormula return Number is begin return PAQUETE.calcular(1); end;"
        source.write_text(original, encoding="utf-8")
        ambiguous = set_function_source(self.project, ["CALCULAR"], "BDATOS", "DL_11G")
        self.assertIn("indique si el prefijo", ambiguous.error)
        self.assertEqual(source.read_text(encoding="utf-8"), original)
        linked = set_function_source(self.project, ["CALCULAR"], "", "DL_11G")
        self.assertIsNone(linked.error)
        self.assertIn("PAQUETE.CALCULAR@DL_11G(1)", source.read_text(encoding="utf-8"))
        unlinked = set_function_source(self.project, ["CALCULAR"])
        self.assertIsNone(unlinked.error)
        self.assertIn("PAQUETE.CALCULAR(1)", source.read_text(encoding="utf-8"))
        self.assertNotIn("return CALCULAR(1)", source.read_text(encoding="utf-8"))
        packaged = set_function_source(self.project, ["CALCULAR"], "BDATOS", "DL_11G", prefix_role="package")
        self.assertIsNone(packaged.error)
        self.assertIn("BDATOS.PAQUETE.CALCULAR@DL_11G(1)", source.read_text(encoding="utf-8"))
        moved = set_function_source(self.project, ["CALCULAR"], "OTRO", "DL_2")
        self.assertIsNone(moved.error)
        self.assertIn("OTRO.PAQUETE.CALCULAR@DL_2(1)", source.read_text(encoding="utf-8"))
        package_unlinked = set_function_source(self.project, ["CALCULAR"])
        self.assertIsNone(package_unlinked.error)
        self.assertIn("OTRO.PAQUETE.CALCULAR(1)", source.read_text(encoding="utf-8"))
        package_plain = set_function_source(self.project, ["CALCULAR"], prefix_role="package")
        self.assertIsNone(package_plain.error)
        self.assertIn("PAQUETE.CALCULAR(1)", source.read_text(encoding="utf-8"))
        self.assertTrue(validate_project(self.project).valid)

    def test_qualified_schema_function_can_be_replaced_explicitly(self) -> None:
        raw = json.loads(self.project.read_text(encoding="utf-8"))
        raw["reports_code"] = "codigo_reports"
        self.project.write_text(json.dumps(raw), encoding="utf-8")
        code_dir = self.generated / "codigo_reports"
        code_dir.mkdir(exist_ok=True)
        source = code_dir / "cf_dobleformula.sql"
        source.write_text("function CF_DOBLEFormula return Number is begin return BDATOS.calcular(1); end;",
                          encoding="utf-8")
        changed = set_function_source(self.project, ["CALCULAR"], "OTRO", "DL_2", prefix_role="schema")
        self.assertIsNone(changed.error)
        self.assertIn("OTRO.CALCULAR@DL_2(1)", source.read_text(encoding="utf-8"))

    def test_add_then_remove_link_keeps_two_package_functions(self) -> None:
        raw = json.loads(self.project.read_text(encoding="utf-8"))
        raw["reports_code"] = "codigo_reports"
        self.project.write_text(json.dumps(raw), encoding="utf-8")
        code_dir = self.generated / "codigo_reports"
        code_dir.mkdir(exist_ok=True)
        source = code_dir / "cf_dobleformula.sql"
        source.write_text(
            "function CF_DOBLEFormula return Number is begin\n"
            "  return P1S_SALDO_CHEQUE1.titular(1) + "
            "P1S_SALDO_CHEQUE1.sum_dscto_garante(1);\nend;\n", encoding="utf-8")
        added = set_function_source(self.project, None, "", "DL_11G")
        self.assertIsNone(added.error)
        self.assertEqual(set(added.functions), {"TITULAR", "SUM_DSCTO_GARANTE"})
        removed = set_function_source(self.project, None)
        self.assertIsNone(removed.error)
        code = source.read_text(encoding="utf-8")
        self.assertIn("P1S_SALDO_CHEQUE1.TITULAR(1)", code)
        self.assertIn("P1S_SALDO_CHEQUE1.SUM_DSCTO_GARANTE(1)", code)
        self.assertNotIn("@DL_11G", code)
        self.assertTrue(validate_project(self.project).valid)

    def test_validation_warns_if_override_loses_package_prefix(self) -> None:
        original = next((Path(__file__).resolve().parents[1] / "proyectos").glob("PSR308Y*"))
        copied = self.root / "PSR308Y"
        shutil.copytree(original, copied)
        project = copied / "generado" / "PSR308Y.report.json"
        override = copied / "generado" / "codigo_reports" / "cf_saldo_actualformula.sql"
        override.write_text(
            "function CF_saldo_actualFormula return Number is begin "
            "return TITULAR(:cod_adm, :nro_che); end;", encoding="utf-8")
        checked = validate_project(project, strict=False)
        self.assertIn("LAYOUT-REPORTS-016", [item.code for item in checked.diagnostics.items])
        restored = restore_function_sources(project, ["TITULAR"])
        self.assertIsNone(restored.error)
        self.assertIn("p1s_saldo_cheque1.titular(:cod_adm, :nro_che)",
                      override.read_text(encoding="utf-8"))
        self.assertEqual(len(restored.backups), 1)
        self.assertIn("return TITULAR(", restored.backups[0].read_text(encoding="utf-8"))
        self.assertNotIn("LAYOUT-REPORTS-016", [item.code for item in
                                               validate_project(project, strict=False).diagnostics.items])

    def test_restore_refuses_ambiguous_call_counts_without_writing(self) -> None:
        original = next((Path(__file__).resolve().parents[1] / "proyectos").glob("PSR308Y*"))
        copied = self.root / "PSR308Y"
        shutil.copytree(original, copied)
        project = copied / "generado" / "PSR308Y.report.json"
        override = copied / "generado" / "codigo_reports" / "cf_saldo_actualformula.sql"
        code = ("function CF_saldo_actualFormula return Number is begin "
                "return TITULAR(:cod_adm, :nro_che) + TITULAR(:cod_adm, :nro_che); end;")
        override.write_text(code, encoding="utf-8")
        change = restore_function_sources(project, ["TITULAR"])
        self.assertIn("2 llamada(s) actuales y 1", change.error)
        self.assertEqual(override.read_text(encoding="utf-8"), code)

    def test_restore_recovers_original_schema_without_replacing_formula_body(self) -> None:
        original = next((Path(__file__).resolve().parents[1] / "proyectos").glob("RPC004*"))
        copied = self.root / "RPC004"
        shutil.copytree(original, copied)
        project = copied / "generado" / "RPC004.report.json"
        override = copied / "generado" / "codigo_reports" / "cf_gradoformula.sql"
        override.parent.mkdir(exist_ok=True)
        override.write_text(
            "function CF_GRADOFormula return Char is begin "
            "return OTRO.F_SACA_DES_GRADO_CORTO@DL_11G(:COD_GRD, :F_DOC); end;",
            encoding="utf-8")
        change = restore_function_sources(project, ["F_SACA_DES_GRADO_CORTO"])
        self.assertIsNone(change.error)
        code = override.read_text(encoding="utf-8")
        self.assertIn("bdatos.f_saca_des_grado_corto(:COD_GRD, :F_DOC)", code)
        self.assertIn("return bdatos.f_saca_des_grado_corto", code)
        self.assertTrue(change.backups[0].exists())


if __name__ == "__main__":
    unittest.main()
