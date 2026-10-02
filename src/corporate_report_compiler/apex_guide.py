"""Instrucciones de publicación en APEX y esqueletos de proyecto.

Este módulo no amplía el contrato: solo describe, a partir de un resultado ya
validado, qué archivo va a cada lugar de APEX, y genera un proyecto inicial
(``.report.json`` + ``.sql``) a partir de los marcadores de un DOCX.
"""

from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
import re

from .compiler import CompilationResult
from .diagnostics import Diagnostics
from .docx_reader import read_template
from .model import TemplateModel
from .project import PROJECT_SCHEMA


PAGE_PLACEHOLDER = "XX"
ITEM_TYPES = {
    "VARCHAR2": "Text Field (o Select List si el filtro usa un LOV)",
    "NUMBER": "Number Field (o Select List: el return value del LOV)",
    "DATE": "Date Picker; misma máscara que format_mask",
    "TIMESTAMP": "Date Picker con hora; misma máscara que format_mask",
}


def _page_of(item: str) -> str:
    match = re.match(r"^P([0-9]+|XX)_", item)
    if match is None:
        return "?"
    page = match.group(1)
    return "0 (Global Page)" if page == "0" else page


def _table(rows: list[tuple[str, ...]]) -> list[str]:
    widths = [max(len(row[index]) for row in rows) for index in range(len(rows[0]))]
    lines = []
    for position, row in enumerate(rows):
        lines.append("  " + " | ".join(value.ljust(widths[index]) for index, value in enumerate(row)).rstrip())
        if position == 0:
            lines.append("  " + "-+-".join("-" * width for width in widths))
    return lines


def page_item_rows(result: CompilationResult) -> list[tuple[str, ...]]:
    """Filas (item, página, tipo APEX, uso, detalle) de los Page Items requeridos."""

    project = result.project
    if project is None:
        return []
    rows: list[tuple[str, ...]] = []
    for binding in project.bindings:
        detail = "obligatorio" if binding["required"] else "opcional"
        if binding.get("format_mask"):
            detail += f"; máscara {binding['format_mask']}"
        rows.append(
            (
                binding["item"],
                _page_of(binding["item"]),
                ITEM_TYPES.get(binding["type"], binding["type"]),
                f"bind :{binding['bind']} ({binding['type']})",
                detail,
            )
        )
    for field in project.fields:
        if field["source"] == "ITEM":
            rows.append(
                (
                    field["item"],
                    _page_of(field["item"]),
                    ITEM_TYPES.get(field["type"], field["type"]),
                    f"{{{{FIELD:{field['name']}}}}} del DOCX",
                    "texto del encabezado o pie",
                )
            )
    if project.format_item:
        rows.append(
            (
                project.format_item,
                _page_of(project.format_item),
                "Select List o Radio Group",
                "formato de salida",
                "valores de retorno exactos: PDF, XLSX (por defecto PDF)",
            )
        )
    if project.orientation_item:
        rows.append(
            (
                project.orientation_item,
                _page_of(project.orientation_item),
                "Select List o Radio Group",
                "orientación A4",
                "valores de retorno exactos: AUTO, PORTRAIT, LANDSCAPE",
            )
        )
    return rows


def build_apex_guide(result: CompilationResult) -> str:
    """Guía paso a paso para publicar un reporte compilado en APEX 24.2."""

    project = result.project
    if not result.valid or project is None or result.definition is None:
        return ""
    process = next((path for path in result.artifacts if path.name == "apex_process.sql"), None)
    request = f"DOWNLOAD_{project.report_id}"[:255]
    lines = [
        "QUÉ SUBIR A APEX Y DÓNDE",
        "=" * 24,
        "",
        "1. Una sola vez por parsing schema: instalar el package común",
        "   Archivo: sql/pkg_corporate_reports.sql (de este proyecto).",
        "   Dónde: SQL Workshop > SQL Scripts > Upload > Run, conectado al parsing",
        "   schema de la aplicación (o @sql/install.sql desde SQLcl/SQL*Plus).",
        "   Verifique: PKG_CORPORATE_REPORTS y su body en estado VALID y sin filas en",
        "   USER_ERRORS. Si ya instaló este mismo archivo (versión 8), omita este paso.",
        "",
        "2. Page Items que debe crear en la página (Page Designer > Items):",
    ]
    rows = page_item_rows(result)
    if rows:
        lines.extend(_table([("Page Item", "Página", "Tipo sugerido", "Uso", "Detalle"), *rows]))
        lines.append("   Active Session State Protection en los items de filtro cuando corresponda.")
    else:
        lines.append("   Ninguno: la consulta no usa binds ni campos ITEM.")
    lines += [
        "",
        "3. Columnas de la plantilla (deben existir como alias en el SELECT):",
    ]
    column_rows = [("Encabezado Word", "Alias SQL", "Ancho", "Impresa")]
    excluded = set(result.definition["excluded_columns"])
    for column in result.definition["columns"]:
        width = column["width"]
        width_text = width["mode"] + (f" {width['value']:g}" if "value" in width else "")
        column_rows.append((column["heading"], column["name"], width_text, "no" if column["name"] in excluded else "sí"))
    lines.extend(_table(column_rows))
    lines += [
        "",
        "4. Botón de descarga (Page Designer > Region > Create Button):",
        f"   Nombre: {request}; Action: Submit Page. No use una Dynamic Action Ajax:",
        "   la descarga necesita un envío completo de la página.",
        "   En los atributos de la página: Advanced > Reload on Submit: Always.",
        "   Con «Only for Success» el navegador espera JSON y la descarga falla.",
        "",
        "5. Proceso de descarga (Page Designer > Processing > Create Process):",
        "   Type: Execute Code; Language: PL/SQL.",
        "   Source > PL/SQL Code: pegue el contenido COMPLETO de:",
        f"     {process if process is not None else 'apex_process.sql (carpeta de salida)'}",
        "   (el bloque declare ... end; tal cual, sin editarlo a mano).",
        "   Execution > Point: Processing.",
        f"   Server-side Condition: When Button Pressed = {request}.",
        "   Security > Authorization Scheme: el mismo que protege la página.",
        "",
        "6. Lo que NO se sube a APEX:",
        "   el .docx, el .report.json, el .sql fuente, template.json y validation.json.",
        "   Se conservan versionados en el repositorio de Informática.",
        "",
        "7. Antes de producción: esta validación es local. Pruebe en APEX 24.2 PDF y XLSX,",
        "   filtros vacíos y obligatorios, cero y muchas filas, y un usuario sin permiso",
        "   (ver Manual de uso, sección de pruebas obligatorias).",
    ]
    return "\n".join(lines)


@dataclass(frozen=True, slots=True)
class ProjectSkeleton:
    project: dict
    sql: str
    mapping: list[tuple[str, ...]]


def _unique(candidate: str, used: set[str], limit: int) -> str:
    value = candidate[:limit]
    counter = 2
    while value in used:
        suffix = f"_{counter}"
        value = candidate[: limit - len(suffix)] + suffix
        counter += 1
    used.add(value)
    return value


def _item_name(page: str, name: str, used: set[str]) -> str:
    return _unique(f"P{page}_{re.sub(r'[^A-Z0-9_]', '_', name.upper())}", used, 128)


def _bind_name(name: str, used: set[str]) -> str:
    # Prefijo F_: evita palabras reservadas de Oracle y nombres reservados de APEX.
    return _unique(f"F_{name.upper()}", used, 30)


def build_skeleton(template: TemplateModel, *, page: str = PAGE_PLACEHOLDER, report_id: str | None = None) -> ProjectSkeleton:
    """Proyecto inicial con un filtro opcional PXX_<ALIAS> por columna del DOCX."""

    page = page.strip().upper() or PAGE_PLACEHOLDER
    stem = template.source_path.stem
    identifier = report_id or re.sub(r"[^A-Z0-9_]", "_", stem.upper()).strip("_")
    if not identifier or not identifier[0].isalpha():
        identifier = f"R_{identifier}"
    identifier = identifier[:30]

    select_lines = []
    filters = []
    bindings = []
    used_items: set[str] = set()
    used_binds: set[str] = set()
    mapping: list[tuple[str, ...]] = [("Marcador DOCX", "Alias / bind SQL", "Page Item", "Uso")]
    for column in template.columns:
        # Alias y columna entre comillas: válidos aunque coincidan con una
        # palabra reservada (DATE, LEVEL, SELECT…). Reemplace t."X" por la
        # columna real de la tabla.
        source_column = f't."{column.name}"'
        select_lines.append(f'       {source_column} as "{column.name}"')
        item = _item_name(page, column.name, used_items)
        bind = _bind_name(column.name, used_binds)
        filters.append(f"(:{bind} is null or {source_column} = :{bind})")
        bindings.append({"bind": bind, "item": item, "type": "VARCHAR2", "required": False})
        mapping.append((f"{{{{COLUMN:{column.name}}}}} ({column.heading})", f":{bind}", item, "filtro opcional"))

    fields = []
    for name in template.fields:
        item = _item_name(page, name, used_items)
        fields.append({"name": name, "source": "ITEM", "item": item, "type": "VARCHAR2"})
        mapping.append((f"{{{{FIELD:{name}}}}}", "-", item, "texto en encabezado/pie"))

    sql_lines = ["select " + select_lines[0].strip() + ("," if len(select_lines) > 1 else "")]
    for index, line in enumerate(select_lines[1:], start=1):
        sql_lines.append(line + ("," if index < len(select_lines) - 1 else ""))
    sql_lines.append("  from tabla_origen t")
    for index, condition in enumerate(filters):
        sql_lines.append((" where " if index == 0 else "   and ") + condition)
    sql = "\n".join(sql_lines) + "\n"

    project = {
        "schema": PROJECT_SCHEMA,
        "report_id": identifier,
        "template": template.source_path.name,
        "query_file": f"{stem}.sql",
        "title": re.sub(r"\s+", " ", stem.replace("_", " ")).strip() or "Reporte",
        "file_name": re.sub(r"[^A-Za-z0-9_-]+", "_", stem).strip("_") or "reporte",
        "orientation": "AUTO",
        "max_rows": 1000,
        "bindings": bindings,
        "fields": fields,
        "excluded_columns": [],
        "column_widths": [],
    }
    return ProjectSkeleton(project=project, sql=sql, mapping=mapping)


def write_skeleton(docx_path: Path, *, page: str = PAGE_PLACEHOLDER) -> tuple[Diagnostics, ProjectSkeleton | None, tuple[Path, ...]]:
    """Lee el DOCX y escribe ``<nombre>.report.json`` y ``<nombre>.sql`` a su lado.

    Nunca sobrescribe archivos existentes: devuelve un diagnóstico si ya están.
    """

    diagnostics = Diagnostics(strict=True)
    template = read_template(Path(docx_path).resolve(), diagnostics)
    if template is None:
        return diagnostics, None, ()
    skeleton = build_skeleton(template, page=page)
    project_path = template.source_path.with_name(f"{template.source_path.stem}.report.json")
    sql_path = template.source_path.with_name(skeleton.project["query_file"])
    existing = [path for path in (project_path, sql_path) if path.exists()]
    if existing:
        for path in existing:
            diagnostics.error("GUIDE-001", "El archivo ya existe; no se sobrescribe.", location=str(path))
        return diagnostics, skeleton, ()
    sql_path.write_text(skeleton.sql, encoding="utf-8", newline="\n")
    project_path.write_text(
        json.dumps(skeleton.project, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
        newline="\n",
    )
    return diagnostics, skeleton, (project_path, sql_path)


def format_skeleton(skeleton: ProjectSkeleton, files: tuple[Path, ...], page: str) -> str:
    lines = ["PROYECTO INICIAL GENERADO DESDE EL DOCX", "=" * 39, ""]
    lines += [f"  {path}" for path in files]
    lines += ["", "Mapeo de marcadores del DOCX a Page Items:"]
    lines += _table(skeleton.mapping)
    lines += [""]
    if page.strip().upper() in {"", PAGE_PLACEHOLDER}:
        lines += [
            f"Los items usan el prefijo P{PAGE_PLACEHOLDER}_: reemplace {PAGE_PLACEHOLDER} por el número de",
            "página APEX en el .report.json (por ejemplo P42_). Hasta entonces la",
            "validación lo informará como Page Item no válido.",
        ]
    lines += [
        "Siguientes pasos:",
        '1. En el .sql, reemplace tabla_origen y las columnas t."ALIAS" por las reales;',
        "   para LOV, haga JOIN y devuelva el display value con el alias del DOCX.",
        "2. Elimine los filtros y bindings que no necesite (deben coincidir 1 a 1).",
        "3. Ajuste título, file_name, column_widths y max_rows en el .report.json.",
        "4. Pulse Validar y luego Compilar.",
    ]
    return "\n".join(lines)
