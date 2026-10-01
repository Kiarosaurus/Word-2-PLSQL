from __future__ import annotations

import json
import math
from pathlib import Path
import re
from typing import Any

from .diagnostics import Diagnostics
from .model import ProjectModel, TemplateModel
from .placeholders import IDENTIFIER_PATTERN, normalize_identifier
from .sql_reader import validate_sql


PROJECT_SCHEMA = "corporate-report-project/1.0"
ITEM_PATTERN = re.compile(r"^P(?:0|[1-9][0-9]*)_[A-Z][A-Z0-9_]{0,119}$")
ALLOWED_TYPES = {"VARCHAR2", "NUMBER", "DATE", "TIMESTAMP"}
ALLOWED_ORIENTATIONS = {"AUTO", "PORTRAIT", "LANDSCAPE"}
ALLOWED_ALIGNMENTS = {"START", "CENTER", "END"}
ALLOWED_FONT_FAMILIES = {"HELVETICA", "TIMES", "COURIER"}
ALLOWED_FONT_WEIGHTS = {"NORMAL", "BOLD"}
COLOR_PATTERN = re.compile(r"^#[0-9A-F]{6}$")
ALLOWED_PROJECT_KEYS = {
    "schema", "report_id", "template", "query_file", "title", "file_name",
    "max_rows", "orientation", "format_item", "orientation_item", "bindings",
    "binds", "fields", "excluded_columns", "exclude_columns", "column_widths",
    "columns", "style_overrides",
}


def _utf8_length(value: str) -> int:
    return len(value.encode("utf-8"))


def _finite_float(value: object) -> float | None:
    try:
        numeric = float(value)  # type: ignore[arg-type]
    except (OverflowError, TypeError, ValueError):
        return None
    return numeric if math.isfinite(numeric) else None


def _validate_varchar_text(
    value: object,
    *,
    maximum_bytes: int,
    label: str,
    code: str,
    diagnostics: Diagnostics,
) -> None:
    if isinstance(value, str) and _utf8_length(value) > maximum_bytes:
        diagnostics.error(
            code,
            f"{label} no puede superar {maximum_bytes} bytes UTF-8.",
        )


def _reject_json_constant(value: str) -> None:
    raise ValueError(f"Constante JSON no permitida: {value}")


def _reject_unknown_keys(
    value: dict[str, Any],
    allowed: set[str],
    *,
    location: str,
    code: str,
    diagnostics: Diagnostics,
) -> None:
    for key in sorted(set(value) - allowed):
        diagnostics.error(code, f"Propiedad no admitida: {key!r}.", location=location)


def _load_json(path: Path, diagnostics: Diagnostics) -> dict[str, Any] | None:
    if not path.is_file():
        diagnostics.error("PROJECT-001", "No existe el archivo de proyecto.", location=str(path))
        return None
    try:
        with path.open("r", encoding="utf-8-sig") as stream:
            value = json.load(stream, parse_constant=_reject_json_constant)
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as exc:
        diagnostics.error("PROJECT-002", f"No se pudo leer el proyecto JSON: {exc}", location=str(path))
        return None
    if not isinstance(value, dict):
        diagnostics.error("PROJECT-003", "La raíz del proyecto debe ser un objeto JSON.", location=str(path))
        return None
    return value


def _safe_relative_file(
    base: Path,
    value: object,
    *,
    label: str,
    diagnostics: Diagnostics,
) -> Path | None:
    if not isinstance(value, str) or not value.strip():
        diagnostics.error("PROJECT-004", f"Falta {label}.")
        return None
    candidate = (base / value).resolve()
    try:
        candidate.relative_to(base.resolve())
    except ValueError:
        diagnostics.error(
            "PROJECT-005",
            f"{label} debe estar dentro de la carpeta del proyecto.",
            location=value,
        )
        return None
    if not candidate.is_file():
        diagnostics.error("PROJECT-006", f"No existe {label}: {value}.", location=value)
        return None
    return candidate


def _validate_bindings(
    raw: object,
    sql_binds: tuple[str, ...],
    diagnostics: Diagnostics,
) -> tuple[dict[str, Any], ...]:
    if raw is None:
        raw = []
    if not isinstance(raw, list):
        diagnostics.error("PROJECT-010", "'bindings' debe ser un array JSON.")
        return ()
    result: list[dict[str, Any]] = []
    seen: set[str] = set()
    for index, entry in enumerate(raw, start=1):
        location = f"bindings[{index}]"
        if not isinstance(entry, dict):
            diagnostics.error("PROJECT-011", "Cada bind debe ser un objeto JSON.", location=location)
            continue
        _reject_unknown_keys(
            entry,
            {"bind", "name", "item", "type", "required", "format_mask", "format"},
            location=location,
            code="PROJECT-094",
            diagnostics=diagnostics,
        )
        raw_name = entry.get("bind", entry.get("name", ""))
        name = normalize_identifier(str(raw_name), label="El bind", diagnostics=diagnostics, location=location)
        if name is None:
            continue
        if name in seen:
            diagnostics.error("PROJECT-012", f"El bind {name} está duplicado.", location=location)
            continue
        seen.add(name)
        item = str(entry.get("item", "")).strip().upper()
        if len(item) > 128 or not ITEM_PATTERN.fullmatch(item):
            diagnostics.error("PROJECT-013", f"El Page Item {item!r} no es válido.", location=location)
        data_type = str(entry.get("type", "VARCHAR2")).strip().upper()
        if data_type not in ALLOWED_TYPES:
            diagnostics.error("PROJECT-014", f"Tipo de bind no admitido: {data_type}.", location=location)
        format_mask = entry.get("format_mask", entry.get("format"))
        if data_type in {"DATE", "TIMESTAMP"} and (not isinstance(format_mask, str) or not format_mask.strip()):
            diagnostics.error("PROJECT-015", f"El bind {name} requiere format_mask.", location=location)
        required = entry.get("required", False)
        if not isinstance(required, bool):
            diagnostics.error("PROJECT-016", "'required' debe ser booleano JSON (true o false).", location=location)
            required = False
        normalized: dict[str, Any] = {
            "bind": name,
            "item": item,
            "type": data_type,
            "required": required,
        }
        if format_mask:
            format_mask = str(format_mask)
            _validate_varchar_text(
                format_mask,
                maximum_bytes=4000,
                label=f"El format_mask de {location}",
                code="PROJECT-098",
                diagnostics=diagnostics,
            )
            normalized["format_mask"] = format_mask
        result.append(normalized)
    sql_set = set(sql_binds)
    for missing in sorted(sql_set - seen):
        diagnostics.error("SQL-020", f"El bind SQL :{missing} no tiene mapeo en 'bindings'.")
    for extra in sorted(seen - sql_set):
        diagnostics.error("SQL-021", f"El mapeo {extra} no corresponde a ningún bind de la consulta.")
    order = {name: position for position, name in enumerate(sql_binds)}
    return tuple(sorted(result, key=lambda item: order.get(item["bind"], len(order))))


def _validate_fields(
    raw: object,
    template: TemplateModel,
    diagnostics: Diagnostics,
) -> tuple[dict[str, Any], ...]:
    if raw is None:
        raw = []
    if not isinstance(raw, list):
        diagnostics.error("PROJECT-030", "'fields' debe ser un array JSON.")
        return ()
    result: list[dict[str, Any]] = []
    seen: set[str] = set()
    for index, entry in enumerate(raw, start=1):
        location = f"fields[{index}]"
        if not isinstance(entry, dict):
            diagnostics.error("PROJECT-031", "Cada campo debe ser un objeto JSON.", location=location)
            continue
        name = normalize_identifier(str(entry.get("name", "")), label="El campo", diagnostics=diagnostics, location=location)
        if name is None:
            continue
        if name in {"REPORT_TITLE", "APP_USER", "GENERATED_AT"}:
            diagnostics.error("PROJECT-032", f"{name} es un campo reservado y no debe mapearse.", location=location)
            continue
        if name in seen:
            diagnostics.error("PROJECT-033", f"El campo {name} está duplicado.", location=location)
            continue
        seen.add(name)
        source = str(entry.get("source", "")).strip().upper()
        if source not in {"ITEM", "CONTEXT", "SYSTEM", "CONSTANT"}:
            diagnostics.error("PROJECT-034", f"Origen de campo no admitido: {source!r}.", location=location)
            continue
        field_keys = {
            "ITEM": {"name", "source", "item", "type", "format_mask", "format"},
            "CONTEXT": {"name", "source", "key"},
            "SYSTEM": {"name", "source", "key", "format"},
            "CONSTANT": {"name", "source", "value"},
        }
        _reject_unknown_keys(
            entry,
            field_keys[source],
            location=location,
            code="PROJECT-095",
            diagnostics=diagnostics,
        )
        normalized: dict[str, Any] = {"name": name, "source": source}
        if source == "ITEM":
            item = str(entry.get("item", "")).strip().upper()
            if len(item) > 128 or not ITEM_PATTERN.fullmatch(item):
                diagnostics.error("PROJECT-035", f"El Page Item {item!r} no es válido.", location=location)
            data_type = str(entry.get("type", "VARCHAR2")).strip().upper()
            if data_type not in ALLOWED_TYPES:
                diagnostics.error("PROJECT-036", f"Tipo de campo no admitido: {data_type}.", location=location)
            normalized.update(item=item, type=data_type)
            format_mask = entry.get("format_mask", entry.get("format"))
            if format_mask:
                format_mask = str(format_mask)
                _validate_varchar_text(
                    format_mask,
                    maximum_bytes=4000,
                    label=f"El format_mask de {location}",
                    code="PROJECT-098",
                    diagnostics=diagnostics,
                )
                normalized["format_mask"] = format_mask
        elif source == "CONTEXT":
            key = str(entry.get("key", "")).strip().upper()
            if key not in {"APP_USER", "APP_ID", "APP_PAGE_ID"}:
                diagnostics.error("PROJECT-037", f"Contexto no admitido: {key!r}.", location=location)
            normalized["key"] = key
        elif source == "SYSTEM":
            key = str(entry.get("key", "")).strip().upper()
            if key != "GENERATED_AT":
                diagnostics.error("PROJECT-038", f"Campo de sistema no admitido: {key!r}.", location=location)
            normalized["key"] = key
            if entry.get("format"):
                system_format = str(entry["format"])
                _validate_varchar_text(
                    system_format,
                    maximum_bytes=4000,
                    label=f"El formato de {location}",
                    code="PROJECT-098",
                    diagnostics=diagnostics,
                )
                normalized["format"] = system_format
        else:
            value = entry.get("value")
            valid_constant = isinstance(value, (str, int, float)) and not isinstance(value, bool)
            if valid_constant and isinstance(value, (int, float)):
                valid_constant = _finite_float(value) is not None
            if not valid_constant:
                diagnostics.error("PROJECT-039", "El valor constante debe ser texto o número.", location=location)
            normalized["value"] = value
            _validate_varchar_text(
                str(value) if valid_constant else value,
                maximum_bytes=4000,
                label=f"La constante de {location}",
                code="PROJECT-099",
                diagnostics=diagnostics,
            )
        result.append(normalized)
    required = set(template.fields)
    for missing in sorted(required - seen):
        diagnostics.error("TOKEN-020", f"El campo {{FIELD:{missing}}} no tiene mapeo.")
    for extra in sorted(seen - required):
        diagnostics.error("TOKEN-021", f"El campo mapeado {extra} no se usa en la plantilla.")

    # APP_USER y GENERATED_AT son placeholders integrados del paquete. No se
    # serializan como campos configurables: el contrato los resuelve aunque el
    # array p_fields_json sea NULL.
    return tuple(result)


def _validate_item_name(raw: object, label: str, diagnostics: Diagnostics) -> str | None:
    if raw is None:
        return None
    value = str(raw).strip().upper()
    if len(value) > 128 or not ITEM_PATTERN.fullmatch(value):
        diagnostics.error("PROJECT-075", f"{label} no es un Page Item válido: {value!r}.")
        return None
    return value


def _validate_style_overrides(raw: object, diagnostics: Diagnostics) -> dict[str, Any]:
    """Validate the optional advanced style override without broadening output.

    Word remains the normal source. This escape hatch is useful when the exact
    native APEX value must be adjusted without editing the DOCX.
    """

    if raw is None:
        return {}
    if not isinstance(raw, dict):
        diagnostics.error("PROJECT-080", "'style_overrides' debe ser un objeto JSON.")
        return {}
    allowed_zones = {"title", "header", "table_header", "table_body", "border", "footer"}
    unknown = sorted(set(raw) - allowed_zones)
    for name in unknown:
        diagnostics.error("PROJECT-081", f"Zona de estilo no admitida: {name!r}.")
    result: dict[str, Any] = {}
    for zone, value in raw.items():
        normalized_zone = "title" if zone == "header" else zone
        if zone not in allowed_zones or not isinstance(value, dict):
            if zone in allowed_zones:
                diagnostics.error("PROJECT-082", f"El estilo {zone} debe ser un objeto.")
            continue
        allowed = {"font_family", "font_size", "font_weight", "font_color", "alignment", "background_color"}
        if normalized_zone == "border":
            allowed = {"width", "color"}
        clean: dict[str, Any] = {}
        for key, item in value.items():
            if key not in allowed:
                diagnostics.error("PROJECT-083", f"Propiedad no admitida en {zone}: {key!r}.")
                continue
            if key in {"font_color", "background_color", "color"}:
                color = str(item).strip().upper()
                if not COLOR_PATTERN.fullmatch(color):
                    diagnostics.error("PROJECT-084", f"Color no válido en {zone}.{key}: {item!r}.")
                clean[key] = color
            elif key == "font_family":
                family = str(item).strip().upper()
                if family not in ALLOWED_FONT_FAMILIES:
                    diagnostics.error("PROJECT-085", f"Fuente no admitida: {family!r}.")
                clean[key] = family
            elif key == "font_weight":
                weight = str(item).strip().upper()
                if weight not in ALLOWED_FONT_WEIGHTS:
                    diagnostics.error("PROJECT-086", f"Peso de fuente no admitido: {weight!r}.")
                clean[key] = weight
            elif key == "alignment":
                alignment = str(item).strip().upper()
                if alignment not in ALLOWED_ALIGNMENTS:
                    diagnostics.error("PROJECT-087", f"Alineación no admitida: {alignment!r}.")
                elif normalized_zone == "title" and alignment != "CENTER":
                    diagnostics.error("PROJECT-093", "El título debe usar alineación CENTER.")
                clean[key] = alignment
            elif key in {"font_size", "width"}:
                if not isinstance(item, (int, float)) or isinstance(item, bool):
                    diagnostics.error("PROJECT-088", f"{zone}.{key} debe ser numérico.")
                else:
                    numeric = _finite_float(item)
                    if numeric is None:
                        diagnostics.error("PROJECT-088", f"{zone}.{key} debe ser un número finito.")
                        continue
                    if key == "width" and not 0 <= numeric <= 5:
                        diagnostics.error("PROJECT-089", "El borde debe estar entre 0 y 5 puntos.")
                    if key == "font_size":
                        limits = {
                            "title": (8, 24),
                            "table_header": (6, 16),
                            "table_body": (6, 14),
                            "footer": (6, 12),
                        }
                        minimum, maximum = limits[normalized_zone]
                        if not minimum <= numeric <= maximum:
                            diagnostics.error(
                                "PROJECT-090",
                                f"El tamaño de fuente de {normalized_zone} debe estar entre "
                                f"{minimum} y {maximum} puntos.",
                            )
                    clean[key] = numeric
        result[normalized_zone] = clean
    return result


def _validate_column_options(
    raw_widths: object,
    raw_overrides: object,
    raw_excluded: object,
    template: TemplateModel,
    diagnostics: Diagnostics,
) -> tuple[tuple[dict[str, Any], ...], tuple[dict[str, Any], ...], tuple[str, ...]]:
    names = [column.name for column in template.columns]
    name_set = set(names)
    excluded: list[str] = []
    if raw_excluded is None:
        raw_excluded = []
    if not isinstance(raw_excluded, list):
        diagnostics.error("PROJECT-040", "'excluded_columns' debe ser un array.")
    else:
        for value in raw_excluded:
            name = str(value).strip().upper()
            if not IDENTIFIER_PATTERN.fullmatch(name):
                diagnostics.error("PROJECT-041", f"Exclusión inválida: {value!r}.")
            elif name not in name_set:
                diagnostics.warning("PROJECT-042", f"La exclusión {name} no coincide con una columna de la plantilla.")
            elif name not in excluded:
                excluded.append(name)
    if name_set and name_set.issubset(set(excluded)):
        diagnostics.error("PROJECT-043", "No se pueden excluir todas las columnas.")

    overrides: list[dict[str, Any]] = []
    if raw_overrides is None:
        raw_overrides = []
    if not isinstance(raw_overrides, list):
        diagnostics.error("PROJECT-044", "'columns' debe ser un array.")
    else:
        seen_overrides: set[str] = set()
        for index, entry in enumerate(raw_overrides, start=1):
            location = f"columns[{index}]"
            if not isinstance(entry, dict):
                diagnostics.error("PROJECT-045", "La configuración de columna debe ser un objeto.", location=location)
                continue
            _reject_unknown_keys(
                entry,
                {"name", "alignment", "format_mask"},
                location=location,
                code="PROJECT-096",
                diagnostics=diagnostics,
            )
            name = str(entry.get("name", "")).strip().upper()
            if name not in name_set:
                diagnostics.error("PROJECT-046", f"La columna {name!r} no existe en la plantilla.", location=location)
                continue
            if name in seen_overrides:
                diagnostics.error("PROJECT-047", f"La columna {name} tiene configuración duplicada.", location=location)
                continue
            seen_overrides.add(name)
            normalized: dict[str, Any] = {"name": name}
            if "alignment" in entry:
                alignment = str(entry["alignment"]).strip().upper()
                if alignment not in ALLOWED_ALIGNMENTS:
                    diagnostics.error("PROJECT-048", f"Alineación no admitida: {alignment}.", location=location)
                normalized["alignment"] = alignment
            if "format_mask" in entry:
                value = entry["format_mask"]
                if value is not None and not isinstance(value, str):
                    diagnostics.error("PROJECT-049", "format_mask debe ser texto o null.", location=location)
                _validate_varchar_text(
                    value,
                    maximum_bytes=4000,
                    label=f"El format_mask de {location}",
                    code="PROJECT-098",
                    diagnostics=diagnostics,
                )
                normalized["format_mask"] = value
            overrides.append(normalized)

    widths: list[dict[str, Any]] = []
    if raw_widths is None:
        raw_widths = []
    if not isinstance(raw_widths, list):
        diagnostics.error("PROJECT-050", "'column_widths' debe ser un array.")
    else:
        seen_widths: set[str] = set()
        auto_count = 0
        fixed_total = 0.0
        for index, entry in enumerate(raw_widths, start=1):
            location = f"column_widths[{index}]"
            if not isinstance(entry, dict):
                diagnostics.error("PROJECT-051", "Cada ancho debe ser un objeto.", location=location)
                continue
            _reject_unknown_keys(
                entry,
                {"column", "mode", "value"},
                location=location,
                code="PROJECT-097",
                diagnostics=diagnostics,
            )
            name = str(entry.get("column", "")).strip().upper()
            if name not in name_set:
                diagnostics.error("PROJECT-052", f"La columna {name!r} no existe en la plantilla.", location=location)
                continue
            if name in seen_widths:
                diagnostics.error("PROJECT-053", f"La columna {name} tiene un ancho duplicado.", location=location)
                continue
            seen_widths.add(name)
            mode = str(entry.get("mode", "")).strip().upper()
            if mode not in {"FIXED_PERCENT", "WEIGHT", "AUTO"}:
                diagnostics.error(
                    "PROJECT-054",
                    f"Modo de ancho no admitido: {mode!r}.",
                    location=location,
                    suggestion="Use FIXED_PERCENT, WEIGHT o AUTO.",
                )
                continue
            normalized: dict[str, Any] = {"column": name, "mode": mode}
            is_visible = name not in excluded
            if mode == "AUTO":
                if is_visible:
                    auto_count += 1
                if "value" in entry:
                    diagnostics.error("PROJECT-055", "AUTO no admite 'value'.", location=location)
            else:
                value = entry.get("value")
                if not isinstance(value, (int, float)) or isinstance(value, bool):
                    diagnostics.error("PROJECT-056", f"{mode} requiere un valor numérico.", location=location)
                    continue
                numeric = _finite_float(value)
                if numeric is None:
                    diagnostics.error("PROJECT-056", f"{mode} requiere un número finito.", location=location)
                    continue
                if mode == "FIXED_PERCENT" and not 1 <= numeric <= 95:
                    diagnostics.error("PROJECT-057", "FIXED_PERCENT debe estar entre 1 y 95.", location=location)
                if mode == "WEIGHT" and numeric <= 0:
                    diagnostics.error("PROJECT-058", "WEIGHT debe ser mayor que cero.", location=location)
                if mode == "FIXED_PERCENT" and is_visible:
                    fixed_total += numeric
                normalized["value"] = numeric
            widths.append(normalized)
        if auto_count > 1:
            diagnostics.error("PROJECT-059", "Solo puede existir una columna AUTO.")
        if fixed_total >= 99:
            diagnostics.error("PROJECT-060", "La suma de FIXED_PERCENT debe ser menor que 99.")

    return tuple(widths), tuple(overrides), tuple(excluded)


def load_project(
    path: Path,
    template: TemplateModel,
    sql: str,
    sql_binds: tuple[str, ...],
    diagnostics: Diagnostics,
) -> ProjectModel | None:
    raw = _load_json(path, diagnostics)
    if raw is None:
        return None
    unknown_keys = sorted(set(raw) - ALLOWED_PROJECT_KEYS - {"paper_size", "auto_reserve_pct"})
    for key in unknown_keys:
        diagnostics.error("PROJECT-092", f"Propiedad de proyecto no admitida: {key!r}.")
    if "paper_size" in raw:
        diagnostics.error(
            "PROJECT-076",
            "'paper_size' no forma parte del contrato: la plantilla DOCX debe estar en A4.",
        )
    if "auto_reserve_pct" in raw:
        diagnostics.error(
            "PROJECT-077",
            "'auto_reserve_pct' no forma parte del contrato; la reserva AUTO es interna al package.",
        )
    schema = raw.get("schema")
    if schema != PROJECT_SCHEMA:
        diagnostics.error("PROJECT-070", f"Esquema no compatible: {schema!r}; se esperaba {PROJECT_SCHEMA!r}.")
    report_id = normalize_identifier(str(raw.get("report_id", "")), label="report_id", diagnostics=diagnostics) or "INVALID"
    title = str(raw.get("title", "")).strip()
    if not title:
        diagnostics.error("PROJECT-071", "El título del reporte no puede estar vacío.")
    elif len(title) > 255:
        diagnostics.error("PROJECT-100", "El título del reporte no puede superar 255 caracteres.")
    file_name = str(raw.get("file_name", report_id.lower())).strip()
    sanitized = re.sub(r"[^A-Za-z0-9_-]+", "_", file_name).strip("_")
    if not sanitized or sanitized != file_name:
        diagnostics.error("PROJECT-072", "file_name solo admite letras, números, _ y -.")
    elif len(file_name) > 180:
        diagnostics.error("PROJECT-101", "file_name no puede superar 180 caracteres.")
    max_rows = raw.get("max_rows", 1000)
    if not isinstance(max_rows, int) or isinstance(max_rows, bool) or not 1 <= max_rows <= 100_000:
        diagnostics.error("PROJECT-073", "max_rows debe ser un entero entre 1 y 100000.")
        max_rows = 1000
    orientation_value = raw.get("orientation")
    orientation = str(orientation_value).strip().upper() if orientation_value is not None else None
    if orientation is not None and orientation not in ALLOWED_ORIENTATIONS:
        diagnostics.error("PROJECT-074", "orientation debe ser AUTO, PORTRAIT o LANDSCAPE.")
        orientation = None

    format_item = _validate_item_name(raw.get("format_item"), "format_item", diagnostics)
    orientation_item = _validate_item_name(raw.get("orientation_item"), "orientation_item", diagnostics)

    raw_bindings = raw.get("bindings")
    if raw_bindings is None and "binds" in raw:
        diagnostics.warning("PROJECT-078", "'binds' está obsoleto; use 'bindings'.")
        raw_bindings = raw.get("binds")
    raw_exclusions = raw.get("excluded_columns")
    if raw_exclusions is None and "exclude_columns" in raw:
        diagnostics.warning("PROJECT-079", "'exclude_columns' está obsoleto; use 'excluded_columns'.")
        raw_exclusions = raw.get("exclude_columns")

    bindings = _validate_bindings(raw_bindings, sql_binds, diagnostics)
    fields = _validate_fields(raw.get("fields"), template, diagnostics)
    widths, overrides, excluded = _validate_column_options(
        raw.get("column_widths"),
        raw.get("columns"),
        raw_exclusions,
        template,
        diagnostics,
    )
    style_overrides = _validate_style_overrides(raw.get("style_overrides"), diagnostics)
    if diagnostics.has_errors:
        return None
    return ProjectModel(
        source_path=path,
        raw=raw,
        report_id=report_id,
        template_path=template.source_path,
        query_path=(path.parent / str(raw["query_file"])).resolve(),
        title=title,
        file_name=file_name,
        max_rows=max_rows,
        orientation=orientation,
        format_item=format_item,
        orientation_item=orientation_item,
        style_overrides=style_overrides,
        bindings=bindings,
        fields=fields,
        excluded_columns=excluded,
        column_widths=widths,
        column_overrides=overrides,
    )


def locate_project_inputs(path: Path, diagnostics: Diagnostics) -> tuple[dict[str, Any], Path | None, Path | None]:
    raw = _load_json(path, diagnostics)
    if raw is None:
        return {}, None, None
    template_path = _safe_relative_file(path.parent, raw.get("template"), label="la plantilla", diagnostics=diagnostics)
    query_path = _safe_relative_file(path.parent, raw.get("query_file"), label="la consulta SQL", diagnostics=diagnostics)
    return raw, template_path, query_path


def read_sql(path: Path, diagnostics: Diagnostics) -> tuple[str, tuple[str, ...]]:
    try:
        sql = path.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeError) as exc:
        diagnostics.error("FILE-020", f"No se pudo leer la consulta SQL: {exc}", location=str(path))
        return "", ()
    sql = sql.strip()
    return sql, validate_sql(sql, diagnostics)
