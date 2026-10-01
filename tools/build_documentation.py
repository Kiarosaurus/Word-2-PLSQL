"""Genera los entregables Word a partir de la documentación Markdown final."""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "deliverables"
NAVY = "283848"
SLATE = "4A5563"
LIGHT = "E8EDF2"
PALE = "F5F7F9"
WHITE = "FFFFFF"
TEXT = RGBColor(37, 43, 49)


def shade(cell, color: str) -> None:
    props = cell._tc.get_or_add_tcPr()
    node = props.find(qn("w:shd"))
    if node is None:
        node = OxmlElement("w:shd")
        props.append(node)
    node.set(qn("w:fill"), color)


def set_cell_margins(cell, top=90, start=110, bottom=90, end=110) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    margins = tc_pr.first_child_found_in("w:tcMar")
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tc_pr.append(margins)
    for edge, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = margins.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            margins.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_repeat_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    node = OxmlElement("w:tblHeader")
    node.set(qn("w:val"), "true")
    tr_pr.append(node)


def prevent_row_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    node = OxmlElement("w:cantSplit")
    node.set(qn("w:val"), "true")
    tr_pr.append(node)


def set_keep(paragraph) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    keep = OxmlElement("w:keepNext")
    p_pr.append(keep)


def add_inline(paragraph, text: str, *, code: bool = False) -> None:
    pattern = re.compile(r"(`[^`]+`|\*\*[^*]+\*\*)")
    for part in pattern.split(text):
        if not part:
            continue
        run = paragraph.add_run(part[1:-1] if part.startswith("`") and part.endswith("`") else
                                part[2:-2] if part.startswith("**") and part.endswith("**") else part)
        if code or (part.startswith("`") and part.endswith("`")):
            run.font.name = "Consolas"
            run.font.size = Pt(8.2)
            run.font.color.rgb = RGBColor(35, 54, 72)
        if part.startswith("**") and part.endswith("**"):
            run.bold = True


def configure_document(document: Document, title: str, subtitle: str) -> None:
    section = document.sections[0]
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.7)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    section.header_distance = Cm(0.8)
    section.footer_distance = Cm(0.8)

    normal = document.styles["Normal"]
    normal.font.name = "Aptos"
    normal.font.size = Pt(9.5)
    normal.font.color.rgb = TEXT
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.line_spacing = 1.08

    for name, size, color in (("Title", 24, NAVY), ("Heading 1", 17, NAVY),
                              ("Heading 2", 13, SLATE), ("Heading 3", 10.5, SLATE)):
        style = document.styles[name]
        style.font.name = "Aptos Display"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.space_before = Pt(10 if name != "Title" else 0)
        style.paragraph_format.space_after = Pt(5)
        style.paragraph_format.keep_with_next = True

    title_p = document.add_paragraph(style="Title")
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.add_run(title)
    sub = document.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run(subtitle)
    r.font.name = "Aptos"
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor.from_string(SLATE)
    sub.paragraph_format.space_after = Pt(18)

    border = document.add_table(rows=1, cols=1)
    border.autofit = True
    shade(border.cell(0, 0), NAVY)
    border.cell(0, 0).height = Cm(0.11)
    document.add_paragraph()

    header = section.header.paragraphs[0]
    header.text = "APEX WORD REPORT COMPILER  ·  DOCUMENTACIÓN 1.0"
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for run in header.runs:
        run.font.name = "Aptos"
        run.font.size = Pt(7.5)
        run.font.color.rgb = RGBColor.from_string(SLATE)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run("Uso interno · Oracle APEX 24.2")
    run.font.name = "Aptos"
    run.font.size = Pt(7.5)
    run.font.color.rgb = RGBColor.from_string(SLATE)


def add_code_block(document: Document, lines: list[str]) -> None:
    table = document.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    shade(cell, PALE)
    set_cell_margins(cell, 100, 140, 100, 140)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.keep_together = True
    for idx, line in enumerate(lines):
        if idx:
            p.add_run("\n")
        add_inline(p, line, code=True)


def add_table(document: Document, rows: list[list[str]]) -> None:
    if not rows:
        return
    width = max(len(row) for row in rows)
    table = document.add_table(rows=len(rows), cols=width)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    set_repeat_header(table.rows[0])
    for row_index, source_row in enumerate(rows):
        prevent_row_split(table.rows[row_index])
        for col_index in range(width):
            cell = table.cell(row_index, col_index)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            shade(cell, NAVY if row_index == 0 else WHITE if row_index % 2 else PALE)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            value = source_row[col_index] if col_index < len(source_row) else ""
            add_inline(p, value)
            for run in p.runs:
                run.font.name = "Aptos"
                run.font.size = Pt(8)
                if row_index == 0:
                    run.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
    document.add_paragraph().paragraph_format.space_after = Pt(1)


def parse_table(lines: list[str], start: int) -> tuple[list[list[str]], int]:
    result: list[list[str]] = []
    index = start
    while index < len(lines) and lines[index].strip().startswith("|"):
        cells = [cell.strip() for cell in lines[index].strip().strip("|").split("|")]
        if not all(re.fullmatch(r":?-{3,}:?", cell.replace(" ", "")) for cell in cells):
            result.append(cells)
        index += 1
    return result, index


def markdown_to_docx(source: Path, destination: Path, title: str, subtitle: str) -> None:
    lines = source.read_text(encoding="utf-8").splitlines()
    document = Document()
    configure_document(document, title, subtitle)
    index = 0
    in_code = False
    code_lines: list[str] = []
    skipped_title = False
    while index < len(lines):
        raw = lines[index]
        stripped = raw.strip()
        if stripped.startswith("```"):
            if in_code:
                add_code_block(document, code_lines)
                code_lines = []
                in_code = False
            else:
                in_code = True
            index += 1
            continue
        if in_code:
            code_lines.append(raw)
            index += 1
            continue
        if not stripped:
            index += 1
            continue
        if stripped.startswith("|"):
            rows, index = parse_table(lines, index)
            add_table(document, rows)
            continue
        heading = re.match(r"^(#{1,4})\s+(.+)$", stripped)
        if heading:
            level = len(heading.group(1))
            text = heading.group(2)
            if not skipped_title and level == 1:
                skipped_title = True
                index += 1
                continue
            paragraph = document.add_paragraph(style=f"Heading {min(max(level - 1, 1), 3)}")
            add_inline(paragraph, text)
            set_keep(paragraph)
            index += 1
            continue
        if stripped.startswith("- "):
            p = document.add_paragraph(style="List Bullet")
            add_inline(p, stripped[2:])
            index += 1
            continue
        numbered = re.match(r"^(\d+)\.\s+(.+)$", stripped)
        if numbered:
            # Número explícito: evita que Word continúe automáticamente una
            # lista anterior situada varias páginas antes.
            p = document.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.65)
            p.paragraph_format.first_line_indent = Cm(-0.45)
            p.add_run(f"{numbered.group(1)}. ")
            add_inline(p, numbered.group(2))
            index += 1
            continue
        p = document.add_paragraph()
        add_inline(p, stripped)
        index += 1

    destination.parent.mkdir(parents=True, exist_ok=True)
    document.core_properties.title = title
    document.core_properties.subject = subtitle
    document.core_properties.author = "Área de Informática"
    document.core_properties.keywords = "Oracle APEX, reportes, Word, APEX_DATA_EXPORT"
    document.save(destination)


def main() -> None:
    markdown_to_docx(
        ROOT / "docs" / "MANUAL.md",
        OUTPUT / "Manual_de_uso.docx",
        "Manual de uso",
        "Compilador local Word → reportes Oracle APEX 24.2",
    )
    markdown_to_docx(
        ROOT / "docs" / "INFORME_ALCANCE.md",
        OUTPUT / "Informe_de_alcance_y_limitaciones.docx",
        "Informe de alcance y limitaciones",
        "Arquitectura, seguridad y fronteras del producto · Versión 1.0",
    )
    for path in sorted(OUTPUT.glob("*.docx")):
        print(path.relative_to(ROOT))


if __name__ == "__main__":
    main()
