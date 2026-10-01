from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from . import __version__
from .diagnostics import Diagnostics
from .model import ProjectModel, TemplateModel


COMPILER_NAME = "apex-word-report-compiler"
SCHEMA_VERSION = "1.0"
_Q_DELIMITERS = (("~", "~"), ("!", "!"), ("^", "^"), ("[", "]"), ("{", "}"), ("(", ")"), ("<", ">"), ("#", "#"), ("/", "/"))


def canonical_json(value: object) -> str:
    return json.dumps(
        value,
        ensure_ascii=False,
        indent=2,
        separators=(",", ": "),
        allow_nan=False,
    ) + "\n"


def compact_json(value: object) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), allow_nan=False)


def _merge_style(base: dict[str, Any], override: dict[str, Any] | None) -> dict[str, Any]:
    result = dict(base)
    if override:
        result.update(override)
    return result


def build_definition(template: TemplateModel, project: ProjectModel, sql: str) -> dict[str, Any]:
    width_by_name = {entry["column"]: dict(entry) for entry in project.column_widths}
    override_by_name = {entry["name"]: entry for entry in project.column_overrides}
    columns: list[dict[str, Any]] = []
    for column in template.columns:
        override = override_by_name.get(column.name, {})
        width = width_by_name.get(
            column.name,
            {"column": column.name, "mode": "WEIGHT", "value": 1.0},
        )
        width_payload = {key: value for key, value in width.items() if key != "column"}
        columns.append(
            {
                "name": column.name,
                "heading": column.heading,
                "alignment": override.get("alignment", column.alignment),
                "format_mask": override.get("format_mask", column.format_mask),
                "width": width_payload,
            }
        )

    styles = {
        "title": _merge_style(template.title_style.to_dict(), project.style_overrides.get("title")),
        "table_header": _merge_style(
            template.table_header_style.to_dict(),
            project.style_overrides.get("table_header"),
        ),
        "table_body": _merge_style(
            template.table_body_style.to_dict(),
            project.style_overrides.get("table_body"),
        ),
        "border": _merge_style(
            {"width": template.border_width, "color": template.border_color},
            project.style_overrides.get("border"),
        ),
        "footer": _merge_style(template.footer_style.to_dict(), project.style_overrides.get("footer")),
    }
    return {
        "schema_version": SCHEMA_VERSION,
        "compiler": {"name": COMPILER_NAME, "version": __version__},
        "source": {
            "file_name": template.source_path.name,
            "sha256": template.source_sha256,
        },
        "report": {
            "id": project.report_id,
            "header_template": template.header_template,
            "title_value": project.title,
            "footer_template": template.footer_template,
            "orientation": project.orientation or template.template_orientation,
            "max_rows": project.max_rows,
            "file_name": project.file_name,
        },
        "query": {"sql": sql, "bindings": list(project.bindings)},
        "fields": list(project.fields),
        "columns": columns,
        "excluded_columns": list(project.excluded_columns),
        "styles": styles,
    }


def _utf8_chunks(value: str, maximum_bytes: int = 8_000) -> list[str]:
    if not value:
        return [""]
    chunks: list[str] = []
    current: list[str] = []
    size = 0
    for character in value:
        encoded_size = len(character.encode("utf-8"))
        if current and size + encoded_size > maximum_bytes:
            chunks.append("".join(current))
            current = []
            size = 0
        current.append(character)
        size += encoded_size
    if current:
        chunks.append("".join(current))
    return chunks


def oracle_q_literal(value: str) -> str:
    for opening, closing in _Q_DELIMITERS:
        terminator = f"{closing}'"
        if terminator not in value:
            return f"q'{opening}{value}{closing}'"
    return "'" + value.replace("'", "''") + "'"


def oracle_expression(value: str, *, clob: bool) -> str:
    chunks = _utf8_chunks(value)
    literals = [oracle_q_literal(chunk) for chunk in chunks]
    if not clob:
        return " ||\n        ".join(literals)
    return " ||\n        ".join(
        [f"to_clob({literals[0]})", *literals[1:]]
    )


def _json_argument(value: object, *, null_if_empty: bool = False) -> str:
    if null_if_empty and value in (None, [], {}):
        return "NULL"
    return oracle_expression(compact_json(value), clob=True)


def build_apex_process(definition: dict[str, Any], project: ProjectModel) -> str:
    report = definition["report"]
    columns = [
        {key: value for key, value in column.items() if key != "width"}
        for column in definition["columns"]
    ]
    widths = [
        {"column": column["name"], **column["width"]}
        for column in definition["columns"]
    ]
    format_expression = f"coalesce(:{project.format_item}, 'PDF')" if project.format_item else "'PDF'"
    orientation_expression = (
        f"coalesce(:{project.orientation_item}, {oracle_q_literal(report['orientation'])})"
        if project.orientation_item
        else oracle_q_literal(report["orientation"])
    )

    return f"""-- GENERADO por {COMPILER_NAME} {__version__}.
-- Revise el SQL y los nombres de Page Items antes de copiar este bloque.
-- Los valores de los items se enlazan en PKG_CORPORATE_REPORTS; no se concatenan.
declare
    l_sql                  varchar2(32767) :=
        {oracle_expression(definition['query']['sql'], clob=False)};
    l_bindings_json        clob := {_json_argument(definition['query']['bindings'], null_if_empty=True)};
    l_fields_json          clob := {_json_argument(definition['fields'], null_if_empty=True)};
    l_columns_json         clob := {_json_argument(columns)};
    l_style_json           clob := {_json_argument(definition['styles'])};
    l_excluded_columns     clob := {_json_argument(definition['excluded_columns'], null_if_empty=True)};
    l_column_widths        clob := {_json_argument(widths)};
begin
    pkg_corporate_reports.download_query(
        p_sql_query             => l_sql,
        p_bindings_json         => l_bindings_json,
        p_fields_json           => l_fields_json,
        p_columns_json          => l_columns_json,
        p_style_json            => l_style_json,
        p_header_template       => {oracle_expression(report['header_template'], clob=False)},
        p_title                 => {oracle_expression(report['title_value'], clob=False)},
        p_footer_template       => {oracle_expression(report['footer_template'], clob=False)},
        p_file_name             => {oracle_expression(report['file_name'], clob=False)},
        p_format                => {format_expression},
        p_orientation           => {orientation_expression},
        p_max_rows              => {int(report['max_rows'])},
        p_excluded_columns_json => l_excluded_columns,
        p_column_widths_json    => l_column_widths
    );
end;
"""


def write_text(path: Path, content: str) -> None:
    path.write_text(content, encoding="utf-8", newline="\n")


def emit_artifacts(
    directory: Path,
    definition: dict[str, Any],
    project: ProjectModel,
    diagnostics: Diagnostics,
) -> tuple[Path, ...]:
    directory.mkdir(parents=True, exist_ok=True)
    definition_path = directory / "template.json"
    process_path = directory / "apex_process.sql"
    validation_path = directory / "validation.json"
    write_text(definition_path, canonical_json(definition))
    write_text(process_path, build_apex_process(definition, project))
    write_text(validation_path, canonical_json(diagnostics.to_dict()))
    return definition_path, validation_path, process_path
