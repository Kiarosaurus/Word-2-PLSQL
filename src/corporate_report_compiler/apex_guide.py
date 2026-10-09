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


def expected_artifact(result: CompilationResult, name: str) -> Path | None:
    """Ruta de un archivo generado: la real si ya se compiló, o la que tendrá al compilar."""

    found = next((path for path in result.artifacts if path.name == name), None)
    if found is not None:
        return found
    from .workspace import workspace_outputs

    folders = workspace_outputs(result.project_path)
    if folders is None:
        return None
    generated, project_dir = folders
    return (project_dir if name == "apex_process.sql" or name.startswith("rpt_") else generated) / name


def process_code(result: CompilationResult) -> str | None:
    """Bloque PL/SQL del proceso APEX, generado en memoria (sirve también tras Validar)."""

    if not result.valid or result.definition is None:
        return None
    if result.kind == "layout":
        from .layout_compiler import build_layout_process

        return build_layout_process(result.definition)
    if result.project is None:
        return None
    from .emitter import build_apex_process

    return build_apex_process(result.definition, result.project)


def _project_items(result: CompilationResult) -> list[dict]:
    """Page Items del .report.json (aunque el proyecto tenga errores)."""

    try:
        raw = json.loads(Path(result.project_path).read_text(encoding="utf-8-sig"))
    except (OSError, ValueError):
        return []
    items = []
    for entry in raw.get("parameters") or raw.get("bindings") or []:
        if isinstance(entry, dict) and entry.get("item"):
            items.append({"item": str(entry["item"]).upper(), "type": str(entry.get("type") or "VARCHAR2").upper(),
                          "bind": str(entry.get("name") or entry.get("bind") or "").upper(),
                          "required": bool(entry.get("required")), "mask": entry.get("format_mask")})
    for field in raw.get("fields") or []:
        if isinstance(field, dict) and str(field.get("source", "")).upper() == "ITEM" and field.get("item"):
            items.append({"item": str(field["item"]).upper(), "type": str(field.get("type") or "VARCHAR2").upper(),
                          "bind": f"{{{{FIELD:{field.get('name')}}}}}", "required": False,
                          "mask": field.get("format_mask")})
    for key, values in (("format_item", "PDF, XLSX"), ("orientation_item", "AUTO, PORTRAIT, LANDSCAPE")):
        if raw.get(key):
            items.append({"item": str(raw[key]).upper(), "type": "LIST", "bind": key, "required": False,
                          "mask": values})
    return items


def build_detailed_instructions(result: CompilationResult) -> str:
    """Instrucciones detalladas para publicar el reporte en APEX 24.2, según el modo."""

    layout = result.kind == "layout"
    definition = result.definition or {}
    report = definition.get("report", {})
    package = (report.get("package") or "rpt_<reporte>").lower()
    process = expected_artifact(result, "apex_process.sql")
    package_file = expected_artifact(result, f"{package}.sql") if layout else None
    items = _project_items(result)
    global_items = [item for item in items if item["item"].startswith("P0_")]
    page_items = [item for item in items if not item["item"].startswith("P0_")]
    pages = sorted({_page_of(item["item"]) for item in page_items}) or ["<su página>"]
    page = pages[0]
    button = "DESCARGAR" if layout else f"DOWNLOAD_{report.get('id') or definition.get('report', {}).get('id', 'REPORTE')}"
    mode = ("LAYOUT (motor PDF en PL/SQL: RPT_PDF + RPT_LAYOUT + package del reporte; solo PDF)" if layout
            else "SIMPLE (APEX_DATA_EXPORT con PKG_CORPORATE_REPORTS; PDF y XLSX)")

    lines = [
        "INSTRUCCIONES DETALLADAS PARA APEX 24.2",
        "=" * 39,
        f"Proyecto: {result.project_path}",
        f"Modo: {mode}",
        "",
    ]
    if not result.valid:
        lines += ["ATENCIÓN: el proyecto tiene errores. Corríjalos y pulse Validar o Compilar; mientras tanto",
                  "estas instrucciones son orientativas.", ""]
    lines += [
        "DÓNDE SE EJECUTA CADA SQL",
        "-" * 25,
        "Todo se ejecuta conectado al PARSING SCHEMA de la aplicación APEX (el esquema dueño de las tablas",
        "que usa el reporte; en APEX: Shared Components > Security Attributes > Parsing Schema).",
        "  - SQL Workshop > SQL Scripts > Upload > (elegir archivo) > Run: recomendado; admite archivos con",
        "    varios bloques terminados en «/». Revise el resultado: «Statements Processed» sin errores.",
        "  - SQL Developer: abra el archivo y use «Run Script» (F5), no «Run Statement» (Ctrl+Enter).",
        "  - SQL*Plus o SQLcl: @ruta\\archivo.sql (los install.sql están pensados para esto).",
        "  - NO use SQL Workshop > SQL Commands: ejecuta una sola sentencia y no sirve para packages.",
        "Después de instalar, verifique en SQL Commands:",
    ]
    objects = "'RPT_PDF','RPT_LAYOUT','" + package.upper() + "'" if layout else "'PKG_CORPORATE_REPORTS'"
    lines += [
        f"  select object_name, object_type, status from user_objects where object_name in ({objects});",
        "  (todos VALID; si alguno queda INVALID: select * from user_errors order by name, sequence;)",
        "",
        "PASO 1. UNA SOLA VEZ POR PARSING SCHEMA: el motor común",
        "-" * 54,
    ]
    if layout:
        lines += [
            "  En SQL Scripts suba y ejecute, EN ESTE ORDEN:",
            "    sql\\modo_layout\\rpt_pdf.pks",
            "    sql\\modo_layout\\rpt_pdf.pkb",
            "    sql\\modo_layout\\rpt_layout.pks",
            "    sql\\modo_layout\\rpt_layout.pkb",
            "  o, en SQL*Plus/SQLcl/SQL Developer (F5): @sql\\modo_layout\\install.sql",
            "  Vuelva a hacerlo solo cuando se actualice la herramienta (cambian esos archivos).",
            "",
            "PASO 2. CADA VEZ QUE COMPILE: el package del reporte",
            "-" * 52,
            f"  Archivo: {package_file or package + '.sql'}",
            "  Súbalo y ejecútelo en SQL Scripts (o F5 en SQL Developer). Debe quedar VALID.",
            "  Si las fórmulas convertidas de Oracle Reports llaman a funciones de la base de datos, esas",
            "  funciones deben existir en el parsing schema (o tener sinónimo y permiso EXECUTE).",
        ]
    else:
        lines += [
            "  Archivo: sql\\modo_simple\\pkg_corporate_reports.sql",
            "  Súbalo y ejecútelo en SQL Scripts (o @sql\\modo_simple\\install.sql en SQL*Plus/SQLcl).",
            "",
            "PASO 2. CADA VEZ QUE COMPILE",
            "-" * 27,
            "  No hay SQL que instalar: todo va dentro del proceso (paso 5).",
        ]
    lines += ["", "PASO 3. PAGE GLOBAL (página 0)", "-" * 30]
    if global_items:
        lines += [
            "  Si la aplicación no tiene Global Page: Create Page > Global Page (queda como página 0).",
            "  En Page Designer de la página 0: Rendering > Body > clic derecho > Create Region",
            "  (Type: Static Content, por ejemplo «Opciones del reporte»). Para que solo aparezca en las",
            f"  páginas de reportes: Server-side Condition > Current Page Is Contained Within Expression 1 = {page}.",
            "  Dentro de esa región cree (clic derecho en la región > Create Page Item):",
        ]
        for item in global_items:
            pairs = ("PDF / PDF, Excel / XLSX" if item["bind"] == "format_item"
                     else "Automática / AUTO, Vertical / PORTRAIT, Horizontal / LANDSCAPE")
            lines += [f"    {item['item']}: Identification > Type: Select List;",
                      f"      List of Values > Type: Static Values; pares Display Value / Return Value: {pairs}",
                      f"      (los Return Value deben ser exactamente {item['mask']})."]
    else:
        lines.append("  No se necesita nada en la página 0" + (" (el modo layout solo genera PDF)." if layout else "."))
    title = f"PASO 4. PÁGINA {page}: los filtros (Page Items)"
    lines += ["", title, "-" * len(title)]
    if page_items:
        lines += [
            f"  En Page Designer de la página {page}: Rendering > Body > clic derecho > Create Region",
            "  (Type: Static Content, por ejemplo «Filtros»). Dentro, clic derecho > Create Page Item:",
        ]
        for item in page_items:
            kind = ITEM_TYPES.get(item["type"], item["type"])
            extra = []
            if item["required"]:
                extra.append("Validation > Value Required: On")
            if item["mask"]:
                extra.append(f"Appearance > Format Mask: {item['mask']}")
            lines.append(f"    {item['item']}  ({kind})  -> {':' + item['bind'] if not item['bind'].startswith('{') else item['bind']}")
            if extra:
                lines.append("      " + "; ".join(extra))
        lines.append("  El nombre del item debe ser EXACTAMENTE el indicado: el proceso lo lee por nombre.")
    else:
        lines.append("  Ninguno: el reporte no tiene filtros.")
    lines += [
        "",
        "PASO 5. BOTÓN Y PROCESO (misma página)",
        "-" * 38,
        f"  Botón: en la región de filtros, clic derecho > Create Button. Button Name: {button};",
        "    Behavior > Action: Submit Page. (No use Dynamic Action ni «Execute Server-side Code».)",
        "  Página: clic en el nombre de la página (raíz del árbol) > Advanced > Reload on Submit: Always.",
        "    Con «Only for Success» la descarga falla (el navegador espera JSON).",
        "  Proceso: pestaña Processing > Processes > clic derecho > Create Process.",
        "    Identification > Type: Execute Code; Source > Location: Local Database; Language: PL/SQL.",
        "    Source > PL/SQL Code: pegue el bloque completo (botón «Copiar código APEX» de esta ventana) o el",
        f"    contenido de {process or 'apex_process.sql'}.",
        f"    Server-side Condition > When Button Pressed: {button}.",
        "    Security > Authorization Scheme: el mismo de la página.",
        "  No agregue Branches después del proceso: la descarga termina la petición.",
        "",
        "PASO 6. PROBAR",
        "-" * 14,
        "  Ejecute la página, complete los filtros y pulse el botón: debe descargarse el "
        + ("PDF." if layout else "PDF (y el XLSX si eligió formato)."),
        "  Pruebe filtros vacíos, uno que no devuelva filas y un usuario sin permiso.",
        "",
        "LO QUE NO SE SUBE A APEX",
        "-" * 24,
        "  El .docx, el .report.json, los q_*.sql / .sql fuente, layout.json o template.json,",
        "  validation.json y el XML de Reports: quedan en la carpeta del proyecto.",
    ]
    return "\n".join(lines)


def build_apex_guide(result: CompilationResult) -> str:
    """Guía paso a paso para publicar un reporte compilado en APEX 24.2."""

    if result.kind == "layout":
        from .layout_compiler import build_layout_guide

        return build_layout_guide(result)
    project = result.project
    if not result.valid or project is None or result.definition is None:
        return ""
    process = expected_artifact(result, "apex_process.sql")
    request = f"DOWNLOAD_{project.report_id}"[:255]
    lines = [
        "QUÉ SUBIR A APEX Y DÓNDE",
        "=" * 24,
        "",
        "1. Una sola vez por parsing schema: instalar el package común",
        "   Archivo: sql/modo_simple/pkg_corporate_reports.sql (de este proyecto).",
        "   Dónde: SQL Workshop > SQL Scripts > Upload > Run, conectado al parsing",
        "   schema de la aplicación (o @sql/modo_simple/install.sql desde SQLcl/SQL*Plus).",
        "   Verifique: PKG_CORPORATE_REPORTS y su body en estado VALID y sin filas en",
        "   USER_ERRORS. Si ya instaló este mismo archivo, omita este paso.",
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
        "3. Ajuste título, file_name y column_widths en el .report.json.",
        "4. Pulse Validar y luego Compilar.",
    ]
    return "\n".join(lines)
