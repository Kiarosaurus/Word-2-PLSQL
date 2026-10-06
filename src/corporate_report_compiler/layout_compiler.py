"""Proyecto y compilación del modo layout (varias tablas, motor PDF en PL/SQL).

Archivo de proyecto ``<nombre>.report.json`` con ``schema`` igual a
``corporate-layout-project/1.0``::

    {
      "schema": "corporate-layout-project/1.0",
      "report_id": "ESTADO_CUENTA",
      "template": "../estado_cuenta.docx",
      "title": "Estado de cuenta del alumno",
      "file_name": "estado_cuenta",
      "parameters": [{"name": "P_COD_ALUMNO", "item": "P71_COD_ALUMNO",
                      "type": "VARCHAR2", "required": true}],
      "queries": {"ALUMNO": "q_alumno.sql", "CUOTAS": "q_cuotas.sql"},
      "constants": {"SISTEMA": "ACADEMIA DEMO"}
    }

Cada consulta es un archivo .sql propio (como una Query del Data Model de
Oracle Reports) y usa los parámetros como binds (``:P_COD_ALUMNO``).
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from pathlib import Path
import re
import shutil
import tempfile
from typing import Any

from . import __version__
from .diagnostics import Diagnostics
from .emitter import COMPILER_NAME, canonical_json, oracle_expression, write_text
from .layout_reader import LayoutTemplate, read_layout_template
from .project import (
    ITEM_PATTERN,
    RESERVED_BIND_NAMES,
    _load_json,
    _safe_relative_file,
    project_root,
    read_sql,
)


LAYOUT_SCHEMA = "corporate-layout-project/1.0"
IDENT = re.compile(r"^[A-Z][A-Z0-9_$#]{0,29}$")
PARAMETER_TYPES = {"VARCHAR2", "NUMBER", "DATE"}
ALLOWED_KEYS = {"schema", "report_id", "template", "title", "file_name", "package", "parameters", "queries", "constants"}
LAYOUT_ARTIFACTS = ("layout.json", "validation.json", "apex_process.sql")


@dataclass(slots=True)
class LayoutProject:
    path: Path
    report_id: str
    template_path: Path
    title: str
    file_name: str
    package: str
    parameters: list[dict[str, Any]]
    queries: dict[str, dict[str, Any]]           # nombre -> {file, path, sql, binds}
    constants: dict[str, str] = field(default_factory=dict)

    @property
    def package_file(self) -> str:
        return f"{self.package.lower()}.sql"


def is_layout_project(path: Path) -> bool:
    try:
        return f'"{LAYOUT_SCHEMA}"' in Path(path).read_text(encoding="utf-8-sig")[:2000]
    except (OSError, UnicodeError):
        return False


def _ident(value: object, label: str, diagnostics: Diagnostics, code: str) -> str | None:
    text = str(value or "").strip().upper()
    if not IDENT.fullmatch(text):
        diagnostics.error(code, f"{label} no es un identificador válido: {value!r}.")
        return None
    return text


def load_layout_project(path: Path, diagnostics: Diagnostics) -> tuple[LayoutProject | None, LayoutTemplate | None]:
    raw = _load_json(path, diagnostics)
    if raw is None:
        return None, None
    if raw.get("schema") != LAYOUT_SCHEMA:
        diagnostics.error("LAYOUT-PROJECT-001", f"'schema' debe ser {LAYOUT_SCHEMA}.")
        return None, None
    for key in sorted(set(raw) - ALLOWED_KEYS):
        diagnostics.error("LAYOUT-PROJECT-002", f"Propiedad desconocida: {key}.")
    root = project_root(path)
    report_id = _ident(raw.get("report_id"), "report_id", diagnostics, "LAYOUT-PROJECT-003")
    template_path = _safe_relative_file(path.parent, raw.get("template"), label="la plantilla",
                                        diagnostics=diagnostics, root=root)
    title = raw.get("title")
    if not isinstance(title, str) or not title.strip() or len(title.encode("utf-8")) > 255:
        diagnostics.error("LAYOUT-PROJECT-004", "'title' es obligatorio y admite hasta 255 bytes UTF-8.")
        title = ""
    file_name = raw.get("file_name") or (report_id or "reporte").lower()
    if not isinstance(file_name, str) or not re.fullmatch(r"[A-Za-z0-9_-]{1,180}", file_name):
        diagnostics.error("LAYOUT-PROJECT-005", "'file_name' admite letras, números, _ o -, hasta 180.")
    package = _ident(raw.get("package") or f"RPT_{report_id or 'REPORTE'}"[:30], "package", diagnostics,
                     "LAYOUT-PROJECT-006")

    parameters: list[dict[str, Any]] = []
    raw_parameters = raw.get("parameters") or []
    if not isinstance(raw_parameters, list):
        diagnostics.error("LAYOUT-PROJECT-010", "'parameters' debe ser un array.")
        raw_parameters = []
    for index, entry in enumerate(raw_parameters, start=1):
        if not isinstance(entry, dict):
            diagnostics.error("LAYOUT-PROJECT-011", f"parameters[{index}] debe ser un objeto.")
            continue
        name = _ident(entry.get("name"), f"parameters[{index}].name", diagnostics, "LAYOUT-PROJECT-012")
        if name in RESERVED_BIND_NAMES:
            diagnostics.error("LAYOUT-PROJECT-013", f"{name} es un nombre reservado de APEX.")
        item = str(entry.get("item") or "").strip().upper()
        if not ITEM_PATTERN.fullmatch(item):
            diagnostics.error("LAYOUT-PROJECT-014", f"parameters[{index}].item no es un Page Item válido: {item!r}.")
        kind = str(entry.get("type") or "VARCHAR2").strip().upper()
        if kind not in PARAMETER_TYPES:
            diagnostics.error("LAYOUT-PROJECT-015", f"Tipo no admitido en parameters[{index}]: {kind}.")
        mask = entry.get("format_mask")
        if kind == "DATE" and not mask:
            diagnostics.error("LAYOUT-PROJECT-016", f"parameters[{index}] es DATE y requiere format_mask.")
        if name:
            parameters.append({"name": name, "item": item, "type": kind, "format_mask": mask,
                               "required": bool(entry.get("required", False))})
    parameter_names = {entry["name"] for entry in parameters}

    queries: dict[str, dict[str, Any]] = {}
    raw_queries = raw.get("queries") or {}
    if not isinstance(raw_queries, dict) or not raw_queries:
        diagnostics.error("LAYOUT-PROJECT-020", "'queries' debe ser un objeto con al menos una consulta.")
        raw_queries = {}
    for raw_name, file in raw_queries.items():
        name = _ident(raw_name, f"La consulta {raw_name!r}", diagnostics, "LAYOUT-PROJECT-021")
        query_path = _safe_relative_file(path.parent, file, label=f"la consulta {raw_name}",
                                         diagnostics=diagnostics, root=root)
        if name is None or query_path is None:
            continue
        query_diagnostics = Diagnostics(strict=diagnostics.strict)
        sql, binds = read_sql(query_path, query_diagnostics)
        for item in query_diagnostics.items:
            location = f"{file}" + (f" — {item.location}" if item.location else "")
            (diagnostics.error if item.severity == "ERROR" else diagnostics.warning)(
                item.code, item.message, location=location)
        for bind in binds:
            if bind not in parameter_names:
                diagnostics.error(
                    "LAYOUT-PROJECT-022",
                    f"La consulta {name} usa :{bind}, que no está declarado en 'parameters'.",
                    location=str(file),
                )
        queries[name] = {"file": str(file), "path": query_path, "sql": sql, "binds": list(binds)}

    constants: dict[str, str] = {}
    raw_constants = raw.get("constants") or {}
    if not isinstance(raw_constants, dict):
        diagnostics.error("LAYOUT-PROJECT-030", "'constants' debe ser un objeto.")
        raw_constants = {}
    for raw_name, value in raw_constants.items():
        name = _ident(raw_name, f"La constante {raw_name!r}", diagnostics, "LAYOUT-PROJECT-031")
        if name in {"REPORT_TITLE", "APP_USER", "GENERATED_AT", "PAGE", "PAGES"} or name in parameter_names:
            diagnostics.error("LAYOUT-PROJECT-032", f"La constante {name} usa un nombre reservado o de parámetro.")
        elif name:
            constants[name] = str(value)

    template = None
    if template_path is not None:
        # Diagnóstico propio: los errores del proyecto no deben ocultar los de la plantilla.
        template_diagnostics = Diagnostics(strict=diagnostics.strict)
        template = read_layout_template(template_path, template_diagnostics)
        for item in template_diagnostics.items:
            (diagnostics.error if item.severity == "ERROR" else diagnostics.warning)(
                item.code, item.message, location=item.location)
    if template is not None:
        for query in sorted(template.queries):
            if query not in queries:
                diagnostics.error("LAYOUT-PROJECT-040", f"La plantilla usa la consulta {query}, que no está en 'queries'.")
        for query in sorted(set(queries) - set(template.queries)):
            diagnostics.warning("LAYOUT-PROJECT-041", f"La consulta {query} no se usa en la plantilla.")
        for name in sorted(template.fields - parameter_names - set(constants)):
            diagnostics.error(
                "LAYOUT-PROJECT-042",
                f"{{{{FIELD:{name}}}}} no es un parámetro ni una constante; use FIELD:CONSULTA.COLUMNA para datos.",
            )
    if diagnostics.has_errors or report_id is None or template_path is None or package is None:
        return None, template
    return LayoutProject(path, report_id, template_path, title, file_name, package, parameters, queries, constants), template


# ------------------------------------------------------------------ emisión
def build_layout_definition(project: LayoutProject, template: LayoutTemplate) -> dict[str, Any]:
    return {
        "schema_version": "layout/1.0",
        "compiler": {"name": COMPILER_NAME, "version": __version__},
        "source": {"file_name": template.source_path.name, "sha256": template.source_sha256},
        "report": {"id": project.report_id, "title": project.title, "file_name": project.file_name,
                   "package": project.package},
        "parameters": project.parameters,
        "queries": {name: {"file": q["file"], "sql": q["sql"], "binds": q["binds"]}
                    for name, q in sorted(project.queries.items())},
        "constants": project.constants,
        "layout": template.layout,
    }


def _literal(value: str | None) -> str:
    return oracle_expression(value, clob=False) if value else "NULL"


def _param_name(name: str) -> str:
    return name.lower() if name.startswith("P_") else f"p_{name.lower()}"


def build_package(definition: dict[str, Any]) -> str:
    from .emitter import compact_json

    report = definition["report"]
    package = report["package"].lower()
    parameters = definition["parameters"]
    signature = ",\n".join(
        f"        {_param_name(p['name'])} IN {p['type']}" for p in parameters
    )
    signature = (signature + ",\n" if signature else "") + "        p_app_user IN VARCHAR2 DEFAULT USER"
    lines = [
        f"-- Generado por {COMPILER_NAME} {__version__} desde {definition['source']['file_name']}.",
        "-- No lo edite a mano: cambie la plantilla Word, las consultas o el proyecto y recompile.",
        f"CREATE OR REPLACE PACKAGE {package} AUTHID CURRENT_USER AS",
        "    FUNCTION build(",
        signature,
        "    ) RETURN BLOB;",
        f"END {package};",
        "/",
        "",
        f"CREATE OR REPLACE PACKAGE BODY {package} AS",
        "",
    ]
    for name, query in definition["queries"].items():
        lines += [
            f"    -- Consulta {name} ({query['file']})",
            f"    FUNCTION q_{name.lower()} RETURN CLOB IS",
            "    BEGIN",
            f"        RETURN {oracle_expression(query['sql'], clob=True)};",
            f"    END q_{name.lower()};",
            "",
        ]
    lines += [
        "    -- Layout compilado desde la plantilla Word.",
        "    FUNCTION layout RETURN CLOB IS",
        "    BEGIN",
        f"        RETURN {oracle_expression(compact_json(definition['layout']), clob=True)};",
        "    END layout;",
        "",
        "    FUNCTION build(",
        signature,
        "    ) RETURN BLOB IS",
        "        l_queries JSON_OBJECT_T := JSON_OBJECT_T();",
        "        l_values  JSON_OBJECT_T := JSON_OBJECT_T();",
        "        l_query   JSON_OBJECT_T;",
        "        l_binds   JSON_ARRAY_T;",
        "    BEGIN",
    ]
    for p in parameters:
        if p["required"]:
            lines += [
                f"        IF {_param_name(p['name'])} IS NULL THEN",
                f"            RAISE_APPLICATION_ERROR(-20740, 'Falta el parámetro obligatorio {p['name']}.');",
                "        END IF;",
            ]
    for name, query in definition["queries"].items():
        lines += [
            "        l_query := JSON_OBJECT_T();",
            f"        l_query.put('sql', q_{name.lower()});",
            "        l_binds := JSON_ARRAY_T();",
            *[f"        l_binds.append('{bind}');" for bind in query["binds"]],
            "        l_query.put('binds', l_binds);",
            f"        l_queries.put('{name}', l_query);",
        ]
    lines.append(f"        l_values.put('REPORT_TITLE', {_literal(report['title'])});")
    lines.append("        l_values.put('APP_USER', p_app_user);")
    for name, value in sorted(definition["constants"].items()):
        lines.append(f"        l_values.put('{name}', {_literal(value)});")
    binds = []
    for p in parameters:
        function = {"VARCHAR2": "bind_text", "NUMBER": "bind_number", "DATE": "bind_date"}[p["type"]]
        binds.append(f"rpt_layout.{function}('{p['name']}', {_param_name(p['name'])})")
    bind_list = ",\n                ".join(binds)
    lines += [
        "        RETURN rpt_layout.render(",
        "            p_layout  => layout,",
        "            p_queries => l_queries.to_clob,",
        "            p_values  => l_values.to_clob,",
        "            p_binds   => rpt_layout.t_binds(" + (f"\n                {bind_list}\n            " if binds else "") + ")",
        "        );",
        "    END build;",
        "",
        f"END {package};",
        "/",
        "",
    ]
    return "\n".join(lines)


def build_layout_process(definition: dict[str, Any]) -> str:
    report = definition["report"]
    arguments = []
    for p in definition["parameters"]:
        item = f":{p['item']}"
        if p["type"] == "NUMBER":
            value = f"to_number({item})"
        elif p["type"] == "DATE":
            value = f"to_date({item}, {_literal(p['format_mask'])})"
        else:
            value = item
        arguments.append(f"        {_param_name(p['name'])} => {value},")
    arguments.append("        p_app_user => :APP_USER")
    file_name = report["file_name"]
    return "\n".join([
        f"-- Proceso APEX generado por {COMPILER_NAME} {__version__} para {report['id']}.",
        "-- Requiere RPT_PDF, RPT_LAYOUT y el package del reporte instalados en el parsing schema.",
        "declare",
        "    l_pdf blob;",
        "begin",
        f"    l_pdf := {report['package'].lower()}.build(",
        *arguments,
        "    );",
        "    sys.htp.init;",
        "    sys.owa_util.mime_header('application/pdf', false);",
        "    sys.htp.p('Content-Length: ' || dbms_lob.getlength(l_pdf));",
        f"    sys.htp.p('Content-Disposition: inline; filename=\"{file_name}.pdf\"');",
        "    sys.owa_util.http_header_close;",
        "    sys.wpg_docload.download_file(l_pdf);",
        "    apex_application.stop_apex_engine;",
        "end;",
        "",
    ])


# ------------------------------------------------------------------ compilación
def validate_layout_project(project_path: Path, *, strict: bool = True):
    from .compiler import CompilationResult

    path = Path(project_path).expanduser().resolve()
    diagnostics = Diagnostics(strict=strict)
    project, template = load_layout_project(path, diagnostics)
    definition = build_layout_definition(project, template) if project and template and not diagnostics.has_errors else None
    return CompilationResult(project_path=path, diagnostics=diagnostics, definition=definition, kind="layout")


def compile_layout_project(project_path: Path, output_directory: Path, *, strict: bool = True,
                           process_directory: Path | None = None):
    from .compiler import _commit_artifacts

    result = validate_layout_project(project_path, strict=strict)
    if not result.valid:
        return result
    definition = result.definition
    output = Path(output_directory).expanduser().resolve()
    upload_dir = Path(process_directory).expanduser().resolve() if process_directory else output
    package_file = f"{definition['report']['package'].lower()}.sql"
    destinations = {"apex_process.sql": upload_dir, package_file: upload_dir}
    sources = {result.project_path.resolve()}
    sources |= {Path(project_path).parent.joinpath(q["file"]).resolve() for q in definition["queries"].values()}
    for name in (*LAYOUT_ARTIFACTS, package_file):
        if (destinations.get(name, output) / name).resolve() in sources:
            result.diagnostics.error("IO-003", f"La salida sobrescribiría un archivo fuente: {name}.")
    if result.diagnostics.has_errors:
        return replace(result, definition=None)
    output.parent.mkdir(parents=True, exist_ok=True)
    staging = Path(tempfile.mkdtemp(prefix=f".{output.name or 'compiled'}-", dir=output.parent))
    try:
        write_text(staging / "layout.json", canonical_json(definition))
        write_text(staging / package_file, build_package(definition))
        write_text(staging / "apex_process.sql", build_layout_process(definition))
        write_text(staging / "validation.json", canonical_json(result.diagnostics.to_dict()))
        staged = tuple(staging / name for name in ("layout.json", "validation.json", package_file, "apex_process.sql"))
        artifacts = _commit_artifacts(staged, output, destinations)
        return replace(result, artifacts=artifacts)
    except OSError as exc:
        result.diagnostics.error("IO-002", f"No se pudieron escribir los artefactos: {exc}", location=str(output))
        return replace(result, definition=None, artifacts=())
    finally:
        shutil.rmtree(staging, ignore_errors=True)


def build_layout_guide(result: Any) -> str:
    definition = result.definition
    if not result.valid or definition is None:
        return ""
    report = definition["report"]
    package_path = next((p for p in result.artifacts if p.name == f"{report['package'].lower()}.sql"), None)
    process_path = next((p for p in result.artifacts if p.name == "apex_process.sql"), None)
    lines = [
        "QUÉ SUBIR A APEX Y DÓNDE (MODO LAYOUT)",
        "=" * 38,
        "",
        "1. Una sola vez por parsing schema: el motor PDF",
        "   sql/modo_layout/install.sql (o, en SQL Workshop, rpt_pdf.pks, rpt_pdf.pkb,",
        "   rpt_layout.pks y rpt_layout.pkb en ese orden). Deben quedar VALID.",
        "",
        "2. Package del reporte (cada vez que recompile):",
        f"   {package_path or report['package'].lower() + '.sql'}",
        "   SQL Workshop > SQL Scripts > Upload > Run. Debe quedar VALID.",
        "",
        "3. Page Items de la página:",
    ]
    for p in definition["parameters"]:
        lines.append(f"   {p['item']:<28} {p['type']}{' (obligatorio)' if p['required'] else ''}  -> :{p['name']}")
    if not definition["parameters"]:
        lines.append("   Ninguno.")
    lines += [
        "",
        "4. Botón Descargar (Action: Submit Page) y en la página Reload on Submit: Always.",
        "5. Proceso en Processing (Execute Code, PL/SQL), When Button Pressed = el botón:",
        f"   pegue {process_path or 'apex_process.sql'}",
        "",
        "Consultas del reporte (cada una en su .sql de generado/):",
    ]
    for name, query in definition["queries"].items():
        binds = ", ".join(f":{b}" for b in query["binds"]) or "sin binds"
        lines.append(f"   {name:<14} {query['file']}  ({binds})")
    return "\n".join(lines)
