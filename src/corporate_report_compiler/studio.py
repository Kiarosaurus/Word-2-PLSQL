"""Local report studio: safe project creation and inspectable sample preview.

The preview consumes the *validated compiler definition*. It deliberately labels
features that require the Oracle PL/SQL renderer instead of claiming PDF parity.
"""
from __future__ import annotations

import html
import json
from pathlib import Path
import re
import tempfile
from typing import Any
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET

from docx import Document
from docx.shared import Mm, Pt, RGBColor

from .compiler import validate_project
from .diagnostics import Diagnostics
from .reports_model import load_reports_model, reports_xml
from .workspace import import_docx


TOKEN = re.compile(r"\{\{(FIELD|COLUMN|SUM|GROUP|GROUP_END):([A-Z][A-Z0-9_$#]*(?:\.[A-Z][A-Z0-9_$#]*)?)(?:\|[^{}]+)?\}\}")
BIND = re.compile(r"(?<!:):([A-Za-z][A-Za-z0-9_$#]*)")


def _save_word(doc: Document, path: Path) -> None:
    """Remove python-docx's stock custom XML parts, forbidden by this compiler."""
    raw = path.with_suffix(".raw.docx")
    doc.save(raw)
    with ZipFile(raw) as source, ZipFile(path, "w", ZIP_DEFLATED) as target:
        for item in source.infolist():
            if item.filename.startswith("customXml/"):
                continue
            content = source.read(item.filename)
            if item.filename == "word/_rels/document.xml.rels":
                content = re.sub(rb'<Relationship\b[^>]*customXml[^>]*/>', b'', content)
            elif item.filename == "[Content_Types].xml":
                content = re.sub(rb'<Override\b[^>]*PartName="/customXml/[^>]*/>', b'', content)
            target.writestr(item, content)
    raw.unlink()


def _word(name: str, queries: dict[str, list[str]], groups: dict[str, list[str]] | None = None,
          summaries: dict[str, list[str]] | None = None) -> Document:
    doc = Document()
    section = doc.sections[0]
    wide = any(len(columns) > 6 for columns in queries.values())
    section.page_width, section.page_height = (Mm(297), Mm(210)) if wide else (Mm(210), Mm(297))
    section.left_margin = section.right_margin = Mm(12)
    doc.styles["Normal"].font.size = Pt(8)
    for style_name in ("Title", "Heading 1"):
        doc.styles[style_name].font.color.rgb = RGBColor(0, 0, 0)
    section.header.paragraphs[0].text = name.replace("_", " ")
    section.footer.paragraphs[0].text = "Página {{PAGE}} de {{PAGES}}"
    title = doc.add_paragraph()
    title_run = title.add_run(name.replace("_", " "))
    title_run.bold = True
    title_run.font.size = Pt(16)
    title_run.font.color.rgb = RGBColor(0, 0, 0)
    for query, columns in queries.items():
        doc.add_heading(query, 1)
        if not columns:
            doc.add_paragraph("Consulta sin columnas identificables; complete el Word antes de compilar.")
            continue
        chosen = columns[:12]
        group_columns = (groups or {}).get(query, [])
        totals = (summaries or {}).get(query, [])
        table = doc.add_table(rows=2 + bool(group_columns) + bool(totals), cols=len(chosen))
        table.style = "Table Grid"
        index = 0
        if group_columns:
            group = group_columns[0]
            table.rows[index].cells[0].text = "Grupo: {{GROUP:%s.%s}}" % (query, group)
            index += 1
        for col, cell in zip(chosen, table.rows[index].cells):
            cell.text = col.replace("_", " ")
        index += 1
        for col, cell in zip(chosen, table.rows[index].cells):
            cell.text = "{{COLUMN:%s.%s}}" % (query, col)
        if totals:
            table.rows[-1].cells[0].text = "Total"
            for col in totals[:min(len(chosen) - 1, 3)]:
                position = min(chosen.index(col) if col in chosen else len(chosen) - 1, len(chosen) - 1)
                table.rows[-1].cells[position].text = "{{SUM:%s.%s}}" % (query, col)
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(6 if "{{" in run.text else 7)
        doc.add_paragraph("")
    return doc


def _sql_columns(sql: str) -> list[str]:
    """Conservative SELECT list extraction; ambiguous expressions get a visible warning."""
    match = re.search(r"\bselect\b(.*?)\bfrom\b", sql, re.I | re.S)
    if not match:
        return []
    source = match.group(1)
    parts, start, depth, quote = [], 0, 0, False
    for i, char in enumerate(source):
        if char == "'":
            quote = not quote
        elif not quote:
            depth += (char == "(") - (char == ")")
            if char == "," and depth == 0:
                parts.append(source[start:i]); start = i + 1
    parts.append(source[start:])
    columns = []
    for part in parts:
        alias = re.search(r"(?:\bas\s+|\s+)([A-Za-z][A-Za-z0-9_$#]*)\s*$", part, re.I)
        bare = re.fullmatch(r"\s*(?:[A-Za-z][\w$#]*\.)?([A-Za-z][\w$#]*)\s*", part)
        name = (alias or bare)
        if name and name.group(1).upper() not in columns:
            columns.append(name.group(1).upper())
    return columns


def create_from_rdf(source: Path, projects_dir: Path, *, page: str = "1") -> dict[str, Any]:
    source, projects_dir = Path(source).resolve(), Path(projects_dir).resolve()
    if source.suffix.lower() not in {".rdf", ".xml"} or not source.is_file():
        raise ValueError("Seleccione un archivo .rdf existente (o .xml para pruebas de regresión).")
    stem = source.stem.removesuffix(".reports")
    target = projects_dir / stem
    if target.exists():
        raise FileExistsError(f"El proyecto ya existe: {target}")
    diagnostics = Diagnostics(strict=True)
    xml = reports_xml(source, diagnostics)
    model = load_reports_model(xml, diagnostics) if xml else None
    if model is None:
        return {"project": None, "diagnostics": diagnostics.to_dict()["diagnostics"]}
    groups: dict[str, list[str]] = {}
    totals: dict[str, list[str]] = {}
    for col in model.columns.values():
        if col.kind == "summary" and col.source and col.source in model.columns:
            totals.setdefault(model.columns[col.source].query, []).append(col.source)
    for link in model.links:
        parent = model.columns.get((link.get("parentColumn") or "").upper())
        if parent:
            groups.setdefault(parent.query, []).append(parent.name)
    unsupported = [f"{q}.{c}: identificador fuera del alfabeto de marcadores Word; agréguelo con alias ASCII en SQL."
                   for q, info in model.queries.items() for c in info["columns"]
                   if c and not re.fullmatch(r"[A-Z][A-Z0-9_$#]{0,29}", c)]
    for q, info in model.queries.items():
        usable = [c for c in info["columns"] if c and re.fullmatch(r"[A-Z][A-Z0-9_$#]{0,29}", c)]
        unsupported += [f"{q}.{c}: columna fuera de las primeras 12 del Word inicial; agréguela en la tabla de detalle."
                        for c in usable[12:]]
    root = ET.parse(xml).getroot()
    for tag in ("repeatingFrame", "frame", "formatTrigger", "anchor", "pageBreak", "pageProtect"):
        elements = list(root.iter(tag))
        if elements:
            names = [e.get("name") for e in elements if e.get("name")]
            unsupported.append(f"{tag}: {len(elements)} elemento(s) en RDF; no se trasladan al Word inicial. "
                               + ("Identificadores: " + ", ".join(names[:15]) + "." if names else ""))
    doc = _word(stem, {q: [c for c in info["columns"] if c and
                                re.fullmatch(r"[A-Z][A-Z0-9_$#]{0,29}", c)]
                             for q, info in model.queries.items()}, groups, totals)
    with tempfile.TemporaryDirectory() as temporary:
        word = Path(temporary) / f"{stem}.docx"
        _save_word(doc, word)
        result = import_docx(word, projects_root=projects_dir, page=page, reports=source)
    notes = unsupported + list(result.notes) + list(result.skipped)
    # The data model omits much of the Reports paper layout. Preserve this as a
    # first-class conversion finding beside the project, never a silent loss.
    notes.append("El Word inicial reconstruye consultas, filas y grupos inferidos; geometría de marcos, anclajes, "
                 "condiciones de formato y saltos del layout RDF requieren revisión manual.")
    if result.paths and result.written:
        structure = {"queries": {q: {"columns": info["columns"], "sql": info["sql"]}
                                 for q, info in model.queries.items()},
                     "groups": model.groups, "links": model.links,
                     "columns": {name: {"query": col.query, "group": col.group, "kind": col.kind,
                                        "source": col.source, "function": col.function, "reset": col.reset,
                                        "formula_code": col.formula_code}
                                 for name, col in model.columns.items()}}
        (result.paths.generated_dir / "studio_model.json").write_text(
            json.dumps(structure, ensure_ascii=False, indent=2), encoding="utf-8")
        (result.paths.generated_dir / "studio_findings.json").write_text(
            json.dumps({"source": str(source), "findings": notes}, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"project": str(result.paths.project) if result.written and result.paths else None,
            "word": str(result.paths.docx) if result.written and result.paths else None,
            "diagnostics": result.diagnostics.to_dict()["diagnostics"], "findings": notes}


def create_from_queries(name: str, query_files: dict[str, Path], projects_dir: Path,
                        *, page: str = "1") -> dict[str, Any]:
    if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{0,29}", name):
        raise ValueError("Nombre de proyecto inválido.")
    target = Path(projects_dir).resolve() / name
    if target.exists():
        raise FileExistsError(f"El proyecto ya existe: {target}")
    sql = {q.upper(): Path(path).read_text(encoding="utf-8-sig") for q, path in query_files.items()}
    columns = {q: _sql_columns(code) for q, code in sql.items()}
    findings = [f"{q}: no se pudo extraer una lista de columnas; defina aliases explícitos." for q, cols in columns.items() if not cols]
    doc = _word(name, columns)
    with tempfile.TemporaryDirectory() as temporary:
        word = Path(temporary) / f"{name}.docx"
        _save_word(doc, word)
        result = import_docx(word, projects_root=projects_dir, page=page)
    if not result.written or not result.paths:
        return {"project": None, "diagnostics": result.diagnostics.to_dict()["diagnostics"], "findings": findings}
    project = result.paths.project
    raw = json.loads(project.read_text(encoding="utf-8"))
    for q, code in sql.items():
        path = result.paths.generated_dir / f"q_{q.lower()}.sql"
        path.write_text(code, encoding="utf-8")
        raw["queries"][q] = path.name
    binds = sorted({bind.upper() for code in sql.values() for bind in BIND.findall(code)})
    raw["parameters"] = [{"name": bind, "item": f"P{page}_{bind}"[:128], "type": "VARCHAR2", "required": False} for bind in binds]
    project.write_text(json.dumps(raw, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return {"project": str(project), "word": str(result.paths.docx), "diagnostics": [], "findings": findings}


def _tokens(value: str, query: str | None, queries: dict[str, Any], sample: dict[str, dict[str, Any]],
            inspector: dict[str, Any], reports: dict[str, Any]) -> str:
    output, previous = [], 0
    for match in TOKEN.finditer(value):
        output.append(html.escape(value[previous:match.start()]))
        kind, name = match.groups()
        owner, _, column = name.partition(".")
        if not column:
            owner, column = query or "", owner
        data = queries.get(owner, {})
        column_model = reports.get("model", {})
        formula = column_model.get("formulas", {}).get(column)
        summary = column_model.get("summaries", {}).get(column)
        key = f"{kind}:{owner}.{column}"
        inspector[key] = {"tipo": kind, "consulta": owner, "columna": column,
                          "sql": data.get("sql", ""), "parámetros": data.get("binds", []),
                          "fórmula": formula, "total": summary,
                          "ámbito": "grupo" if kind.startswith("GROUP") else ("consulta" if kind == "SUM" else "fila de detalle"),
                          "repetición": "una vez por cambio de grupo" if kind.startswith("GROUP") else
                          ("una vez al cerrar la tabla" if kind == "SUM" else "una vez por fila de la consulta"),
                          "sin_filas": "no se imprime la fila de detalle", "salto_de_página": "la cabecera de tabla puede repetirse; verifique en Oracle"}
        actual = sample.get(owner, {}).get(column)
        if kind == "SUM":
            actual = "Σ 42.00" if summary or column else "—"
        output.append(f'<span class="token inspectable" tabindex="0" data-key="{html.escape(key)}">{html.escape(str(actual if actual is not None else "∅"))}</span>')
        previous = match.end()
    output.append(html.escape(value[previous:]))
    return "".join(output)


def _paragraph(node: dict[str, Any], query: str | None, queries: dict[str, Any],
               sample: dict[str, dict[str, Any]], inspector: dict[str, Any], reports: dict[str, Any]) -> str:
    lines = node.get("ln", [])
    return "<p>" + "<br>".join(_tokens("".join(fragment.get("t", "") for fragment in line), query, queries, sample, inspector, reports)
                                for line in lines) + "</p>"


def _blocks(nodes: list[dict[str, Any]], queries: dict[str, Any], sample: dict[str, dict[str, Any]],
            inspector: dict[str, Any], reports: dict[str, Any]) -> str:
    out = []
    for node in nodes:
        kind = node.get("k")
        if kind == "P":
            out.append(_paragraph(node, None, queries, sample, inspector, reports))
        elif kind in {"D", "G"}:
            query = node.get("q")
            table_key = f"TABLE:{query or 'LAYOUT'}:{len(inspector)}"
            info = queries.get(query or "", {})
            inspector[table_key] = {"tipo": "tabla de datos" if kind == "D" else "tabla de maquetación",
                                    "consulta": query, "sql": info.get("sql", ""), "parámetros": info.get("binds", []),
                                    "repetición": "una fila por registro" if kind == "D" else "una vez por bloque",
                                    "ámbito": "consulta" if kind == "D" else "reporte",
                                    "sin_filas": "no imprime detalle" if kind == "D" else "se mantiene",
                                    "salto_de_página": "el motor Oracle decide el corte y repite cabeceras configuradas",
                                    "filtros_registro_RDF": node.get("rf", []),
                                    "filtros_grupo_RDF": node.get("gt", []),
                                    "previsualización": "Los filtros basados en fórmulas PL/SQL se aplican al PDF en Oracle; los datos ficticios no los evalúan."}
            out.append(f'<section class="table inspectable" data-key="{html.escape(table_key)}" data-query="{html.escape(query or "")}">')
            out.append(f'<h3>{html.escape(query or "Tabla")}</h3><table>')
            rows = [(row, 0) for row in node.get("hd", [])]
            if node.get("bd"):
                rows += [(node["bd"], i) for i in range(1, 7)]
            rows += [(row, 0) for row in node.get("ft", [])]
            if not rows:
                rows = [(row, 0) for row in node.get("rows", [])]
            for row, number in rows:
                if number and query:
                    names = list(sample.get(query, {}))
                    sample[query] = {name: (None if number == 4 and i % 3 == 0 else
                                            (round(12.5 * number * (i + 1), 2) if re.search(r"SALDO|MONTO|TOTAL|IMPORTE|IMP_|DOL_", name)
                                             else f"{name} {1 if number < 4 else 2}-{number}"))
                                     for i, name in enumerate(names)}
                out.append("<tr>")
                for cell in row.get("cs", []):
                    out.append("<td>" + "".join(_paragraph(p, query, queries, sample, inspector, reports) for p in cell.get("ps", [])) + "</td>")
                out.append("</tr>")
            out.append("</table></section>")
    return "".join(out)


def preview_project(project: Path, output: Path) -> dict[str, Any]:
    result = validate_project(project, strict=False)
    definition = result.definition
    if definition is None:
        return {"valid": False, "diagnostics": result.diagnostics.to_dict()["diagnostics"]}
    queries = definition.get("queries", {})
    reports = definition.get("reports", {})
    sample: dict[str, dict[str, Any]] = {}
    for query, info in queries.items():
        names = info.get("columns") or _sql_columns(info.get("sql", ""))
        sample[query] = {name: (None if i % 5 == 4 else (12.5 * (i + 1) if re.search(r"SALDO|MONTO|TOTAL|IMPORTE", name) else f"{name} 1"))
                         for i, name in enumerate(names)}
    inspector: dict[str, Any] = {}
    layout = definition.get("layout", {})
    body = _blocks(layout.get("body", []), queries, sample, inspector, reports)
    header = _blocks(layout.get("header", []), queries, sample, inspector, reports)
    footer = _blocks(layout.get("footer", []), queries, sample, inspector, reports)
    if "query" in definition:
        info = definition["query"]
        columns = definition.get("columns", [])
        labels = "".join("<th>" + html.escape(col.get("heading", col["name"])) + "</th>" for col in columns)
        rows = []
        for number in range(1, 7):
            cells = []
            for index, col in enumerate(columns):
                name = col["name"]
                key = f"COLUMN:MAIN.{name}"
                inspector[key] = {"tipo": "COLUMN", "consulta": "MAIN", "columna": name,
                                  "sql": info["sql"], "parámetros": info.get("bindings", []),
                                  "ámbito": "fila de detalle", "repetición": "una vez por fila SQL",
                                  "sin_filas": "tabla sin filas", "salto_de_página": "APEX_DATA_EXPORT decide el salto"}
                value = "∅" if number == 4 and index % 3 == 0 else f"{name} {number}"
                cells.append(f'<td><span class="token inspectable" data-key="{html.escape(key)}">{html.escape(value)}</span></td>')
            rows.append("<tr>" + "".join(cells) + "</tr>")
        header = "<h1>" + html.escape(definition["report"].get("title_value", "Reporte")) + "</h1>"
        body = "<table><tr>" + labels + "</tr>" + "".join(rows) + "</table>"
        footer = html.escape(definition["report"].get("footer_template", ""))
    diagnostics = result.diagnostics.to_dict()["diagnostics"]
    findings_path = Path(project).parent / "studio_findings.json"
    findings = json.loads(findings_path.read_text(encoding="utf-8")).get("findings", []) if findings_path.exists() else []
    structure_path = Path(project).parent / "studio_model.json"
    if structure_path.exists():
        structure = json.loads(structure_path.read_text(encoding="utf-8"))
        for key, item in inspector.items():
            column = structure.get("columns", {}).get(item.get("columna"), {})
            item["grupo"] = column.get("group")
            item["fórmula_original"] = column.get("formula_code")
            item["reinicio_total"] = column.get("reset")
            item["relaciones"] = [link for link in structure.get("links", [])
                                 if link.get("childQuery", "").upper() == item.get("consulta")]
    payload = {"inspector": inspector, "diagnostics": diagnostics, "findings": findings}
    script = "const model=" + json.dumps(payload, ensure_ascii=False).replace("</", "<\\/") + ";" + """
document.querySelectorAll('.inspectable').forEach(el=>el.addEventListener('click',event=>{
 event.stopPropagation();
 const item=model.inspector[el.dataset.key]; document.querySelector('#inspector').textContent=JSON.stringify(item,null,2);
 document.querySelectorAll('.active').forEach(x=>x.classList.remove('active'));el.classList.add('active');
}));
document.querySelector('#diagnostics').textContent=JSON.stringify({diagnósticos:model.diagnostics,hallazgos:model.findings},null,2);
"""
    page = f"""<!doctype html><html lang="es"><meta charset="utf-8"><title>Vista previa {html.escape(Path(project).stem)}</title>
<style>body{{font:14px Arial;margin:0;background:#e7eaf0;display:grid;grid-template-columns:minmax(0,1fr) 390px;height:100vh}}
main{{overflow:auto;padding:24px}}aside{{overflow:auto;background:#fff;padding:18px;border-left:1px solid #bbc}}
.paper{{background:white;max-width:850px;min-height:1100px;margin:auto;padding:55px;box-shadow:0 2px 12px #aaa;box-sizing:border-box}}
table{{border-collapse:collapse;width:100%;font-size:10px}}td{{border:1px solid #bbb;padding:3px;vertical-align:top}}td p{{margin:0}}
.table{{overflow-x:auto}}.token{{background:#d9edff;cursor:pointer;border-bottom:1px solid #3985bb}}.active{{background:#ffda78}}
pre{{white-space:pre-wrap;overflow-wrap:anywhere;font-size:12px}}.note{{background:#fff4d5;padding:10px;margin-bottom:16px}}</style>
<main><div class="note">Datos ficticios deterministas. Vista estructural del layout compilado; el ajuste de líneas, fórmulas PL/SQL y saltos reales requieren RPT_LAYOUT en Oracle/APEX.</div>
<div class="paper">{header}{body}{footer}</div></main><aside><h2>Inspector</h2><p>Seleccione un dato azul.</p><pre id="inspector"></pre><h2>Diagnósticos</h2><pre id="diagnostics"></pre></aside>
<script>{script}</script></html>"""
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(page, encoding="utf-8")
    return {"valid": result.valid, "preview": str(output), "elements": len(inspector), "diagnostics": diagnostics, "findings": findings}
