from __future__ import annotations

from pathlib import Path
import shutil
import subprocess
import re
from collections import Counter

from corporate_report_compiler.compiler import compile_project, validate_project
from corporate_report_compiler.studio import create_from_queries, create_from_rdf, preview_project
from corporate_report_compiler.reports_model import find_rwconverter


ROOT = Path(__file__).resolve().parents[1]


def projects():
    return sorted((ROOT / "proyectos").rglob("*.report.json"))


def test_corpus_validation_compilation_and_preview(tmp_path):
    found = projects()
    assert len(found) >= 8
    for project in found:
        name = project.parent.parent.name
        checked = validate_project(project, strict=False)
        assert checked.valid, (name, checked.diagnostics.to_dict())
        built = compile_project(project, tmp_path / "compiled" / name, strict=False)
        assert built.valid and built.artifacts, (name, built.diagnostics.to_dict())
        viewed = preview_project(project, tmp_path / "preview" / f"{name}.html")
        assert viewed["valid"] and viewed["elements"] > 0, name
        assert Path(viewed["preview"]).is_file()


def test_xml_import_round_trip(tmp_path):
    sources = sorted((ROOT / "proyectos").rglob("*.reports.xml"))
    assert any("FCR108_B" in str(p) for p in sources)
    assert any("FCR118_E" in str(p) for p in sources)
    for xml in sources:
        result = create_from_rdf(xml, tmp_path / "projects")
        assert result["project"], (xml, result)
        checked = validate_project(Path(result["project"]), strict=False)
        assert checked.valid, (xml, checked.diagnostics.to_dict())
        viewed = preview_project(Path(result["project"]), tmp_path / "preview" / f"{xml.parent.parent.name}.html")
        assert viewed["valid"] and viewed["elements"] > 0
        assert result["findings"]


def test_query_to_word_to_preview(tmp_path):
    sql = tmp_path / "orders.sql"
    sql.write_text("select order_id, amount as total_amount from orders where customer_id = :CUSTOMER_ID", encoding="utf-8")
    created = create_from_queries("ORDERS", {"ORDERS": sql}, tmp_path / "projects")
    assert created["project"]
    project = Path(created["project"])
    assert validate_project(project, strict=False).valid
    viewed = preview_project(project, tmp_path / "orders.html")
    assert viewed["valid"] and viewed["elements"] >= 2


def test_reference_pdf_is_available_and_readable():
    original = ROOT / "proyectos/demo/estado_cuenta/ejemplo_estado_cuenta.pdf"
    assert original.read_bytes().startswith(b"%PDF-")
    pdfinfo = shutil.which("pdfinfo")
    if pdfinfo:
        details = subprocess.run([pdfinfo, str(original)], capture_output=True, text=True, check=True).stdout
        assert "Pages:" in details and "Page size:" in details


def test_deploy_reference_pdf_structure_when_available():
    archive = ROOT.parent / "Migración"
    pdfinfo, pdftotext = shutil.which("pdfinfo"), shutil.which("pdftotext")
    if not archive.is_dir() or not pdfinfo or not pdftotext:
        return
    originals = [path for path in archive.rglob("*") if path.suffix.lower() == ".pdf"
                 and any(code in path.name.upper() for code in ("PFR201_3", "PSR308W_1I", "PSR308Y", "RPC004", "FCR118_E"))
                 and "emitido" not in path.name.lower()]
    assert len(originals) >= 5
    for path in originals:
        metadata = subprocess.run([pdfinfo, str(path)], capture_output=True, text=True, check=True).stdout
        pages = re.search(r"^Pages:\s*(\d+)", metadata, re.M)
        assert pages and int(pages.group(1)) >= 1, path
        extracted = subprocess.run([pdftotext, "-f", "1", "-l", "1", str(path), "-"],
                                   capture_output=True, text=True, errors="replace", check=True).stdout
        assert len(extracted.strip()) > 100, path
    original = next(path for path in originals if path.name.lower() == "fcr118_e - original.pdf")
    emitted = next(path for path in archive.rglob("*") if path.name.lower() == "fcr118_e - emitido.pdf")
    def numbers(path):
        text = subprocess.run([pdftotext, str(path), "-"], capture_output=True,
                              text=True, errors="replace", check=True).stdout
        return Counter(re.findall(r"(?<!\w)[\d,]+\.\d{2}(?!\w)", text))
    first, second = numbers(original), numbers(emitted)
    assert sum((first & second).values()) / min(sum(first.values()), sum(second.values())) > 0.70


def test_real_rdf_import_when_converter_and_archive_available(tmp_path):
    archive = ROOT.parent / "Migración" / "22-10-2026" / "FCR118_E.RDF"
    if not archive.is_file() or find_rwconverter() is None:
        return
    imported = create_from_rdf(archive, tmp_path / "projects", page="70")
    assert imported["project"], imported
    assert validate_project(Path(imported["project"]), strict=False).valid
    assert any("repeatingFrame" in finding for finding in imported["findings"])
