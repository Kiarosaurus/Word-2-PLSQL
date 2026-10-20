"""Tablas de las consultas y fórmulas de un proyecto (modo layout) y su origen.

El SQL copiado de Oracle Reports nombra las tablas como las veía el usuario
de Reports (``PGPAGO``). En APEX el parsing schema puede necesitarlas con su
esquema y su DB link (``BDATOS.PGPAGO@DL_11G``). Este módulo encuentra cada
tabla de los FROM/JOIN de:

- ``generado/q_*.sql`` (consultas del reporte), y
- las funciones de Oracle Reports que usan las fórmulas; al cambiarlas se
  escribe su versión corregida en ``generado/codigo_reports/<función>.sql``,
  que reemplaza a la del .rdf al compilar.

y la reescribe, una tabla a la vez, con el esquema y el DB link elegidos.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import json
from pathlib import Path
import re

from .diagnostics import Diagnostics
from .reports_convert import CODE_DIR, tokens
from .reports_model import load_reports_model


# Palabras que terminan la lista de tablas de un FROM (o que no pueden ser un alias).
CLAUSE_END = set("""
where group order connect start union intersect minus having join left right inner full cross natural
on using partition for window model fetch offset into values set returning pivot unpivot sample
when then else end loop is as select and or not
""".split())
# FROM dentro de estas funciones no es el de una consulta: EXTRACT(YEAR FROM f), TRIM(LEADING '0' FROM c).
FROM_FUNCTIONS = {"extract", "trim"}
IGNORED_TABLES = {"DUAL"}
NAME = re.compile(r"^[A-Za-z][A-Za-z0-9_$#]{0,127}$")
LINK = re.compile(r"^[A-Za-z][A-Za-z0-9_$#.]{0,127}$")


@dataclass(frozen=True, slots=True)
class TableSpan:
    start: int
    end: int
    text: str          # tal como está escrito: PGPAGO, bdatos.pgpago@DL_11G...
    table: str         # solo el nombre de la tabla, en mayúsculas


@dataclass(slots=True)
class TableUse:
    label: str         # q_principal.sql, fórmula cf_finalformula
    text: str          # forma actual


@dataclass(slots=True)
class ProjectTable:
    table: str
    uses: list[TableUse] = field(default_factory=list)

    @property
    def forms(self) -> dict[str, int]:
        """Formas actuales y cuántas veces aparece cada una."""

        counts: dict[str, int] = {}
        for use in self.uses:
            counts[use.text] = counts.get(use.text, 0) + 1
        return counts

    @property
    def places(self) -> list[str]:
        return list(dict.fromkeys(use.label for use in self.uses))


def _with_offsets(code: str) -> list[tuple[str, str, int]]:
    items, position = [], 0
    for kind, text in tokens(code):
        items.append((kind, text, position))
        position += len(text)
    return items


def table_spans(code: str) -> list[TableSpan]:
    """Tablas de los FROM y JOIN de un SQL o PL/SQL (sin subconsultas, funciones de tabla ni DUAL)."""

    items = [item for item in _with_offsets(code) if item[0] not in {"ws", "comment"}]
    found: list[TableSpan] = []
    openers: list[str] = []          # identificador antes de cada «(» abierto

    def word(index: int) -> str:
        return items[index][1].lower() if 0 <= index < len(items) and items[index][0] == "ident" else ""

    def text(index: int) -> str:
        return items[index][1] if 0 <= index < len(items) else ""

    def table_ref(index: int) -> int:
        """Lee una tabla desde ``index``; devuelve el índice siguiente (tras el alias)."""

        if items[index][0] != "ident" or word(index) in CLAUSE_END:
            return index
        start, last, table = index, index, items[index][1]
        while text(last + 1) == "." and last + 2 < len(items) and items[last + 2][0] == "ident":
            last += 2
            table = items[last][1]
        if text(last + 1) == "@" and last + 2 < len(items) and items[last + 2][0] == "ident":
            last += 2
            while text(last + 1) == "." and last + 2 < len(items) and items[last + 2][0] == "ident":
                last += 2
        if text(last + 1) == "(":    # TABLE(...) u otra función de tabla
            return last + 1
        end = items[last][2] + len(items[last][1])
        if table.upper() not in IGNORED_TABLES:
            found.append(TableSpan(items[start][2], end, code[items[start][2]:end], table.upper()))
        following = last + 1
        if word(following) and word(following) not in CLAUSE_END:     # alias
            following += 1
        return following

    index = 0
    while index < len(items):
        kind, value, _ = items[index]
        if value == "(":
            openers.append(word(index - 1))
        elif value == ")" and openers:
            openers.pop()
        lowered = value.lower() if kind == "ident" else ""
        if lowered == "from" and not (openers and openers[-1] in FROM_FUNCTIONS):
            following = table_ref(index + 1)
            while text(following) == ",":
                following = table_ref(following + 1)
            index = max(index + 1, following)
            continue
        if lowered == "join":
            index = max(index + 1, table_ref(index + 1))
            continue
        index += 1
    return found


def qualified(table: str, schema: str = "", link: str = "") -> str:
    schema, link = schema.strip().upper(), link.strip().upper()
    return (f"{schema}." if schema else "") + table.upper() + (f"@{link}" if link else "")


# ------------------------------------------------------------------ proyecto
@dataclass(slots=True)
class _Source:
    label: str
    path: Path                 # archivo que se reescribe
    code: str                  # código actual (de ese archivo o del .rdf)
    unit: str | None = None    # función de Reports, si es una fórmula


def _project_parts(project_path: Path) -> tuple[dict, Path] | None:
    try:
        raw = json.loads(Path(project_path).read_text(encoding="utf-8-sig"))
    except (OSError, ValueError):
        return None
    return (raw, Path(project_path).parent) if isinstance(raw, dict) else None


def _used_units(project_path: Path) -> set[str] | None:
    """Funciones de Reports que el package usa (None si no se puede saber).

    Se leen del layout.json de la última compilación (rápido); si no hay, se valida el proyecto.
    """

    compiled = Path(project_path).parent / "layout.json"
    try:
        reports = json.loads(compiled.read_text(encoding="utf-8")).get("reports")
        if reports is not None:
            return {unit["name"] for unit in reports["units"]}
    except (OSError, ValueError, AttributeError, KeyError, TypeError):
        pass
    from .compiler import validate_project

    result = validate_project(Path(project_path))
    reports = (result.definition or {}).get("reports")
    return {unit["name"] for unit in reports["units"]} if reports else None


def _sources(project_path: Path) -> list[_Source]:
    parts = _project_parts(project_path)
    if parts is None:
        return []
    raw, folder = parts
    sources: list[_Source] = []
    files = list((raw.get("queries") or {}).values()) + ([raw["query_file"]] if raw.get("query_file") else [])
    for name in sorted(map(str, files)):
        path = folder / str(name)
        try:
            sources.append(_Source(path.name, path, path.read_text(encoding="utf-8-sig")))
        except OSError:
            continue
    if raw.get("reports_model"):
        model = load_reports_model(folder / str(raw["reports_model"]), Diagnostics())
        if model is not None:
            code_dir = folder / str(raw.get("reports_code") or CODE_DIR)
            used = _used_units(project_path)
            for unit in sorted(model.program_units):
                override = code_dir / f"{unit}.sql"
                if used is not None and unit not in used and not override.exists():
                    continue
                code = override.read_text(encoding="utf-8-sig") if override.exists() else model.program_units[unit]
                sources.append(_Source(f"fórmula {unit}", override, code, unit))
    return sources


def project_tables(project_path: Path) -> list[ProjectTable]:
    """Tablas que usan las consultas y fórmulas del proyecto, con dónde y cómo aparecen."""

    tables: dict[str, ProjectTable] = {}
    for source in _sources(project_path):
        for span in table_spans(source.code):
            tables.setdefault(span.table, ProjectTable(span.table)).uses.append(TableUse(source.label, span.text))
    return [tables[name] for name in sorted(tables)]


@dataclass(slots=True)
class SourceChange:
    tables: list[str] = field(default_factory=list)
    files: list[tuple[Path, int]] = field(default_factory=list)   # archivo, reemplazos
    error: str | None = None


def _formula_file(unit: str, code: str) -> str:
    return "\n".join([
        f"-- {unit}: función de Oracle Reports con el origen de sus tablas ajustado",
        "-- («Administrar items y tablas»). Este archivo reemplaza a la función del .rdf al compilar.",
        "-- Sintaxis de Reports: :NOMBRE para columnas, parámetros y fórmulas.",
        "",
        code.strip(),
        "",
    ])


def set_table_source(project_path: Path, tables: list[str] | None, schema: str = "", link: str = "") -> SourceChange:
    """Escribe ``ESQUEMA.TABLA@DBLINK`` (o solo TABLA si ambos están vacíos) en cada uso de ``tables``.

    ``tables`` None = todas las tablas del proyecto.
    """

    change = SourceChange()
    schema, link = schema.strip().upper(), link.strip().upper()
    if schema and not NAME.fullmatch(schema):
        change.error = f"«{schema}» no es un nombre de esquema válido (letras, números, _, $ o #)."
        return change
    if link and not LINK.fullmatch(link):
        change.error = f"«{link}» no es un nombre de DB link válido (letras, números, _, $, # o puntos)."
        return change
    wanted = {table.upper() for table in tables} if tables is not None else None
    project = Path(project_path)
    parts = _project_parts(project)
    if parts is None:
        change.error = "No se pudo leer el .report.json."
        return change
    raw, _folder = parts
    wrote_formula = False
    for source in _sources(project):
        spans = [span for span in table_spans(source.code) if wanted is None or span.table in wanted]
        new_code, count, changed = source.code, 0, []
        for span in sorted(spans, key=lambda item: item.start, reverse=True):
            target = qualified(span.table, schema, link)
            if span.text != target:
                new_code = new_code[:span.start] + target + new_code[span.end:]
                count += 1
                changed.append(span.table)
        if not count:
            continue
        try:
            if source.unit is None:
                source.path.write_text(new_code, encoding="utf-8", newline="\n")
            else:
                source.path.parent.mkdir(parents=True, exist_ok=True)
                content = new_code if source.path.exists() else _formula_file(source.unit, new_code)
                source.path.write_text(content, encoding="utf-8", newline="\n")
                wrote_formula = True
        except OSError as exc:
            change.error = f"No se pudo escribir {source.path}: {exc}"
            return change
        change.files.append((source.path, count))
        change.tables.extend(changed)
    if wrote_formula and not raw.get("reports_code"):
        raw["reports_code"] = CODE_DIR
        project.write_text(json.dumps(raw, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n")
    change.tables = sorted(set(change.tables))
    if not change.files and change.error is None:
        change.error = "No hay cambios: esas tablas ya tienen ese origen."
    return change


def common_source(project_path: Path, tables: list[ProjectTable] | None = None) -> tuple[str, str]:
    """(esquema, DB link) que ya usa la mayoría de las tablas calificadas, para proponerlos."""

    counts: dict[tuple[str, str], int] = {}
    for table in tables if tables is not None else project_tables(project_path):
        for use in table.uses:
            match = re.fullmatch(r"(?:(\w[\w$#]*)\.)?\w[\w$#]*(?:@([\w$#.]+))?", use.text)
            if match and (match.group(1) or match.group(2)):
                key = ((match.group(1) or "").upper(), (match.group(2) or "").upper())
                counts[key] = counts.get(key, 0) + 1
    return max(counts, key=counts.get) if counts else ("", "")
