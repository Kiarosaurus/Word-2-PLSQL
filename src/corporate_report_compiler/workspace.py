"""Organización de los reportes en ``proyectos/<nombre>/``.

Cada plantilla Word tiene una carpeta con su mismo nombre::

    proyectos/
      <nombre>/
        <nombre>.docx          plantilla (copia del DOCX cargado)
        apex_process.sql       lo único que se pega en APEX
        generado/
          <nombre>.sql         consulta fuente (se incrusta en apex_process.sql)
          <nombre>.report.json proyecto: binds, campos, anchos, estilos
          template.json        definición compilada
          validation.json      diagnósticos de la compilación

Cargar de nuevo un DOCX con el mismo nombre solo reemplaza esos archivos; cada
uno que existía se conserva como ``generado/<archivo>.bak``. Cualquier otro
archivo de la carpeta (documentación, notas) no se toca.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import json
import os
from pathlib import Path
import re
import shutil

from .apex_guide import PAGE_PLACEHOLDER, ProjectSkeleton, _table, build_skeleton
from .diagnostics import Diagnostics
from .docx_reader import read_template
from .layout_reader import is_layout_docx, read_layout_template
from .reports_convert import CODE_DIR as REPORTS_CODE_DIR, plan_conversion
from .reports_model import LEGACY_TEMP_PREFIX, TEMP_DIR, discard_conversion, load_reports_model, reports_xml
from .project import GENERATED_DIRNAME


PROJECTS_DIRNAME = "proyectos"
PROCESS_NAME = "apex_process.sql"
GENERATED_ARTIFACTS = ("template.json", "validation.json")


def default_projects_root() -> Path:
    """``proyectos/`` en la raíz de esta instalación (junto a ``src/``)."""

    return Path(__file__).resolve().parents[2] / PROJECTS_DIRNAME


def is_workspace_project(project_path: Path) -> bool:
    return Path(project_path).parent.name.casefold() == GENERATED_DIRNAME


def workspace_outputs(project_path: Path) -> tuple[Path, Path] | None:
    """(carpeta de template.json/validation.json, carpeta de apex_process.sql)."""

    path = Path(project_path)
    if not is_workspace_project(path):
        return None
    return path.parent, path.parent.parent


PLACEHOLDER_ITEM = re.compile(rf"\bP{PAGE_PLACEHOLDER}_[A-Za-z0-9_$#]+", re.IGNORECASE)
PAGE_ITEM = re.compile(r"\bP([1-9][0-9]{0,5})_[A-Za-z0-9_$#]+")
SCANNED_SUFFIXES = {".json", ".sql", ".txt"}


def page_placeholder_locations(project_dir: Path) -> list[tuple[Path, int, list[str]]]:
    """(archivo, línea, items) de cada Page Item P{XX}_ que falta numerar en la carpeta del proyecto.

    Se omiten las copias .bak y los archivos generados (template.json,
    layout.json, validation.json): Compilar los rehace desde el .report.json.
    """

    found = []
    folder = Path(project_dir)
    if not folder.is_dir():
        return found
    skip = {*GENERATED_ARTIFACTS, "layout.json"}
    for path in sorted(folder.rglob("*")):
        if not path.is_file() or path.suffix.lower() not in SCANNED_SUFFIXES or path.name in skip:
            continue
        try:
            text = path.read_text(encoding="utf-8-sig", errors="replace")
        except OSError:
            continue
        for number, line in enumerate(text.splitlines(), 1):
            items = list(dict.fromkeys(match.upper() for match in PLACEHOLDER_ITEM.findall(line)))
            if items:
                found.append((path, number, items))
    return found


def existing_page(docx_path: Path, projects_root: Path | None = None) -> str | None:
    """Página APEX que ya usa el proyecto de este DOCX (la más frecuente en sus Page Items), o None."""

    source = Path(docx_path).resolve()
    project = WorkspacePaths.for_name((projects_root or default_projects_root()).resolve(), source.stem).project
    try:
        text = project.read_text(encoding="utf-8-sig")
    except OSError:
        return None
    pages = [page for page in PAGE_ITEM.findall(text) if page != "0"]
    return max(sorted(set(pages)), key=pages.count) if pages else None


@dataclass(frozen=True, slots=True)
class WorkspacePaths:
    project_dir: Path
    generated_dir: Path
    docx: Path
    sql: Path
    project: Path

    @classmethod
    def for_name(cls, root: Path, name: str) -> "WorkspacePaths":
        project_dir = root / name
        generated = project_dir / GENERATED_DIRNAME
        return cls(
            project_dir=project_dir,
            generated_dir=generated,
            docx=project_dir / f"{name}.docx",
            sql=generated / f"{name}.sql",
            project=generated / f"{name}.report.json",
        )

    def managed_files(self) -> tuple[Path, ...]:
        """Archivos que la carga reemplaza (con copia .bak). Nada más se toca."""

        return (
            self.docx,
            self.project_dir / PROCESS_NAME,
            self.sql,
            self.project,
            *(self.generated_dir / name for name in GENERATED_ARTIFACTS),
        )


@dataclass(slots=True)
class ImportResult:
    diagnostics: Diagnostics
    paths: WorkspacePaths | None = None
    skeleton: ProjectSkeleton | None = None
    written: tuple[Path, ...] = ()
    backups: tuple[Path, ...] = field(default_factory=tuple)
    layout: bool = False
    notes: tuple[str, ...] = ()
    skipped: tuple[str, ...] = ()      # consultas del Modelo de Datos que no se copiaron, con el motivo


def _same_file(first: Path, second: Path) -> bool:
    try:
        return first.exists() and second.exists() and os.path.samefile(first, second)
    except OSError:
        return False


def _layout_identifier(stem: str) -> str:
    identifier = re.sub(r"[^A-Z0-9_]", "_", stem.upper()).strip("_") or "REPORTE"
    if not identifier[0].isalpha():
        identifier = f"R_{identifier}"
    return identifier[:26]


def _layout_managed(paths: WorkspacePaths, queries: set[str]) -> tuple[Path, ...]:
    identifier = _layout_identifier(paths.docx.stem)
    return (
        paths.docx,
        paths.project_dir / PROCESS_NAME,
        paths.project_dir / f"rpt_{identifier.lower()}.sql",
        paths.project,
        paths.generated_dir / "layout.json",
        paths.generated_dir / "validation.json",
        paths.generated_dir / f"{paths.docx.stem}.reports.xml",
        *(paths.generated_dir / f"q_{query.lower()}.sql" for query in sorted(queries)),
    )


def _reports_model(reports: Path | None, diagnostics: Diagnostics):
    """(modelo, xml) del Oracle Reports original, o (None, None)."""

    if reports is None:
        return None, None
    xml_path = reports_xml(Path(reports).expanduser().resolve(), diagnostics)
    model = load_reports_model(xml_path, diagnostics) if xml_path is not None else None
    return model, xml_path


def _project_queries(template, model) -> tuple[set[str], object]:
    """Consultas del proyecto: las del Word y las que necesitan las fórmulas, totales y Data Links."""

    if model is None:
        return set(template.queries), None
    plan = plan_conversion(model, {column for columns in template.queries.values() for column in columns})
    return set(template.queries) | {query for query in plan.queries if query in model.queries}, plan


def files_to_replace(docx_path: Path, projects_root: Path | None = None, reports: Path | None = None) -> tuple[Path, ...]:
    """Archivos existentes que una carga de ``docx_path`` reemplazaría."""

    source = Path(docx_path).resolve()
    paths = WorkspacePaths.for_name((projects_root or default_projects_root()).resolve(), source.stem)
    managed = paths.managed_files()
    if reports is not None or is_layout_docx(source):
        model, _xml = _reports_model(reports, Diagnostics(strict=True))
        template = read_layout_template(source, Diagnostics(strict=True), model)
        managed = _layout_managed(paths, _project_queries(template, model)[0] if template else set())
    return tuple(
        path
        for path in managed
        if path.exists() and not (path == paths.docx and _same_file(path, source))
    )


def _layout_query_skeleton(name: str, columns: set[str]) -> str:
    lines = [
        f"-- Consulta {name}: pegue aquí el SQL de la Query del reporte original y conserve",
        "-- como alias los nombres de columna que usa la plantilla. Parámetros: :P_NOMBRE.",
    ]
    ordered = sorted(columns)
    select = [f'       t."{column}" as "{column}"' for column in ordered] or ["       1 as dummy"]
    select[0] = "select " + select[0].lstrip()
    lines += [line + ("," if index < len(select) - 1 else "") for index, line in enumerate(select)]
    lines.append("  from tabla_origen t")
    return "\n".join(lines) + "\n"


def _skipped_queries(model, query_names: set[str]) -> list[str]:
    """Consultas del .rdf que el proyecto no copia y por qué (el reporte no las necesita)."""

    lines = []
    for name in sorted(set(model.queries) - set(query_names)):
        columns = {column.name for column in model.columns.values() if column.query == name}
        shown = sorted(field for field, (source, _mask) in model.fields.items() if source in columns)
        if shown:
            reason = ("el Word no usa ninguna de sus columnas, pero el diseño del .rdf SÍ las muestra en "
                      + ", ".join(shown[:6]) + (" ..." if len(shown) > 6 else "")
                      + ": agregue esos marcadores al Word si los necesita")
        else:
            reason = ("ni el Word, ni las fórmulas, totales o Data Links que usa la necesitan; tampoco "
                      "tiene campos en el diseño del .rdf (no se imprime en el reporte original)")
        lines.append(f"{model.queries[name]['name']}: {reason}.")
    return lines


def _commented(text: str) -> list[str]:
    return ["--   " + line.rstrip() for line in text.strip().splitlines()]


def _reports_query_file(name: str, model, plan) -> str:
    """q_<consulta>.sql con el SQL real de Oracle Reports y lo que el motor hace con él."""

    query = model.queries[name]
    lines = [
        f"-- Consulta {query['name']} copiada del Modelo de Datos de Oracle Reports ({model.name}).",
        "-- Los nombres de columna de Reports se asignan por posición: NO cambie el orden del SELECT.",
        "-- Para agregar columnas, agréguelas AL FINAL del SELECT con su alias.",
    ]
    formulas = plan.query_formulas.get(name, [])
    if formulas:
        lines += ["--", "-- Fórmulas de este grupo (convertidas al package del reporte; se calculan por fila):",
                  *(f"--   {formula}" for formula in formulas)]
    for item in plan.filters.get(name, []):
        lines += ["--", f"-- Data Link de Reports: cada fila se filtra con {item['c']} {item['op']} {item['p']} "
                        "(columna del grupo padre). Lo hace el motor: no lo agregue al SQL."]
    for note in plan.pending:
        if f"q_{name.lower()}.sql" in note:
            lines += ["--", f"-- PENDIENTE: {note}"]
    return "\n".join(lines) + "\n" + query["sql"].rstrip() + "\n"


def _reports_code_file(unit, model) -> str:
    """Función de Reports que no se pudo convertir: el usuario la corrige aquí (se convierte al compilar)."""

    owner = next((c.name for c in model.columns.values() if c.unit == unit.name), unit.name)
    return "\n".join([
        f"-- {owner}: función de Oracle Reports que no se pudo convertir sola.",
        f"-- Motivo: {unit.stub}.",
        "-- Corríjala aquí con la sintaxis de Reports (:NOMBRE para columnas, parámetros y fórmulas)",
        "-- y vuelva a compilar: este archivo reemplaza a la función del .rdf. Si la deja así,",
        "-- la fórmula devuelve NULL.",
        "",
        model.program_units.get(unit.name, "").strip(),
        "",
    ])


PARAMETER_TYPES_FROM_REPORTS = {"number": "NUMBER", "date": "DATE"}


def _import_layout_docx(source: Path, root: Path, *, replace: bool, page: str = PAGE_PLACEHOLDER,
                        reports: Path | None = None) -> ImportResult:
    diagnostics = Diagnostics(strict=True)
    model, xml_path = _reports_model(reports, diagnostics)
    if reports is not None and model is None:
        return ImportResult(diagnostics)
    template = read_layout_template(source, diagnostics, model)
    if template is None:
        return ImportResult(diagnostics)
    paths = WorkspacePaths.for_name(root, source.stem)
    existing = files_to_replace(source, root, reports)
    if existing and not replace:
        for path in existing:
            diagnostics.error(
                "GUIDE-002",
                "Ya existe; vuelva a cargar con reemplazo para sustituirlo (se guarda una copia .bak).",
                location=str(path),
            )
        return ImportResult(diagnostics, paths=paths)
    identifier = _layout_identifier(source.stem)
    query_names, plan = _project_queries(template, model)
    notes: list[str] = []
    skipped = _skipped_queries(model, query_names) if model is not None else []
    files: dict[str, str] = {}
    code_files: dict[str, str] = {}
    parameter_names: set[str] = set()
    for name in sorted(query_names):
        if model is not None and name in model.queries:
            files[name] = _reports_query_file(name, model, plan)
            for bind in re.findall(r":([A-Za-z][A-Za-z0-9_$#]*)", model.queries[name]["sql"]):
                if bind.upper() in model.parameters:
                    parameter_names.add(bind.upper())
        else:
            files[name] = _layout_query_skeleton(name, template.queries.get(name, set()))
    if plan is not None:
        parameter_names |= plan.parameters
        for unit in plan.units:
            if unit.stub:
                code_files[unit.name] = _reports_code_file(unit, model)
        notes += plan.pending
        if plan.externals:
            notes.append("Funciones de la base de datos que usan las fórmulas (deben existir en el parsing schema): "
                         + ", ".join(sorted(plan.externals)) + ".")
    constants = {}
    for name in sorted(template.fields):
        if model is not None and name in model.parameters:
            parameter_names.add(name)
        else:
            constants[name] = name
    page_label = page.strip().upper() or PAGE_PLACEHOLDER
    parameters = []
    for name in sorted(parameter_names):
        kind = PARAMETER_TYPES_FROM_REPORTS.get((model.parameters[name]["datatype"] or "").lower(), "VARCHAR2")
        item_name = name[2:] if name.startswith("P_") and len(name) > 2 else name
        entry = {"name": name, "item": f"P{page_label}_{item_name}"[:128], "type": kind, "required": False}
        if kind == "DATE":
            entry["format_mask"] = "DD/MM/YYYY"
        parameters.append(entry)
    project = {
        "schema": "corporate-layout-project/1.0",
        "report_id": identifier,
        "template": f"../{paths.docx.name}",
        "title": re.sub(r"\s+", " ", source.stem.replace("_", " ")).strip() or "Reporte",
        "file_name": re.sub(r"[^A-Za-z0-9_-]+", "_", source.stem).strip("_") or "reporte",
        **({"reports_model": f"{paths.docx.stem}.reports.xml"} if model is not None else {}),
        **({"reports_code": REPORTS_CODE_DIR} if code_files else {}),
        "parameters": parameters,
        "queries": {name: f"q_{name.lower()}.sql" for name in sorted(query_names)},
        "constants": constants,
    }
    backups: list[Path] = []
    written: list[Path] = []
    try:
        paths.generated_dir.mkdir(parents=True, exist_ok=True)
        for path in existing:
            backup = paths.generated_dir / f"{path.name}.bak"
            os.replace(path, backup)
            backups.append(backup)
        if not _same_file(paths.docx, source):
            shutil.copyfile(source, paths.docx)
        written.append(paths.docx)
        if xml_path is not None:
            target = paths.generated_dir / f"{paths.docx.stem}.reports.xml"
            if not _same_file(xml_path, target):
                shutil.copyfile(xml_path, target)
            written.append(target)
        for name, content in files.items():
            query_path = paths.generated_dir / f"q_{name.lower()}.sql"
            query_path.write_text(content, encoding="utf-8", newline="\n")
            written.append(query_path)
        for name, content in code_files.items():
            # Las correcciones del usuario se conservan al volver a cargar el Word.
            code_path = paths.generated_dir / REPORTS_CODE_DIR / f"{name}.sql"
            if not code_path.exists():
                code_path.parent.mkdir(parents=True, exist_ok=True)
                code_path.write_text(content, encoding="utf-8", newline="\n")
                written.append(code_path)
        paths.project.write_text(json.dumps(project, ensure_ascii=False, indent=2) + "\n",
                                 encoding="utf-8", newline="\n")
        written.append(paths.project)
    except OSError as exc:
        diagnostics.error("IO-010", f"No se pudo preparar la carpeta del proyecto: {exc}", location=str(paths.project_dir))
        return ImportResult(diagnostics, paths=paths, backups=tuple(backups))
    return ImportResult(diagnostics, paths=paths, written=tuple(written), backups=tuple(backups), layout=True,
                        notes=tuple(notes), skipped=tuple(skipped))


def import_docx(
    docx_path: str | Path,
    *,
    projects_root: str | Path | None = None,
    page: str = PAGE_PLACEHOLDER,
    replace: bool = False,
    reports: str | Path | None = None,
) -> ImportResult:
    """Copia el DOCX a ``proyectos/<nombre>/`` y genera el proyecto inicial.

    El DOCX se valida antes de tocar ningún archivo. Si la carpeta ya tiene
    archivos de una carga anterior, se exige ``replace=True``; entonces cada
    archivo reemplazado se mueve antes a ``generado/<archivo>.bak``. Con
    ``reports`` (el .rdf o .xml del Oracle Reports original), los marcadores se
    resuelven con su Modelo de Datos y cada consulta se crea con su SQL real.
    """

    diagnostics = Diagnostics(strict=True)
    source = Path(docx_path).expanduser().resolve()
    root = Path(projects_root or default_projects_root()).expanduser().resolve()
    if reports is not None or is_layout_docx(source):
        try:
            return _import_layout_docx(source, root, replace=replace, page=page,
                                       reports=Path(reports) if reports is not None else None)
        finally:
            # El XML ya quedó en generado/<nombre>.reports.xml: la conversión temporal sobra.
            discard_conversion(Path(reports).expanduser() if reports is not None else None)
    template = read_template(source, diagnostics)
    if template is None:
        return ImportResult(diagnostics)

    paths = WorkspacePaths.for_name(root, source.stem)
    existing = files_to_replace(source, root)
    if existing and not replace:
        for path in existing:
            diagnostics.error(
                "GUIDE-002",
                "Ya existe; vuelva a cargar con reemplazo para sustituirlo (se guarda una copia .bak).",
                location=str(path),
            )
        return ImportResult(diagnostics, paths=paths)

    skeleton = build_skeleton(template, page=page)
    skeleton.project["template"] = f"../{paths.docx.name}"
    skeleton.project["query_file"] = paths.sql.name

    backups: list[Path] = []
    try:
        paths.generated_dir.mkdir(parents=True, exist_ok=True)
        for path in existing:
            backup = paths.generated_dir / f"{path.name}.bak"
            os.replace(path, backup)
            backups.append(backup)
        if not _same_file(paths.docx, source):
            shutil.copyfile(source, paths.docx)
        paths.sql.write_text(skeleton.sql, encoding="utf-8", newline="\n")
        paths.project.write_text(
            json.dumps(skeleton.project, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
            newline="\n",
        )
    except OSError as exc:
        diagnostics.error("IO-010", f"No se pudo preparar la carpeta del proyecto: {exc}", location=str(paths.project_dir))
        return ImportResult(diagnostics, paths=paths, skeleton=skeleton, backups=tuple(backups))
    return ImportResult(
        diagnostics,
        paths=paths,
        skeleton=skeleton,
        written=(paths.docx, paths.sql, paths.project),
        backups=tuple(backups),
    )


def format_import(result: ImportResult, page: str) -> str:
    paths = result.paths
    assert paths is not None
    if result.layout:
        lines = ["PROYECTO (MODO LAYOUT) PREPARADO DESDE EL DOCX", "=" * 46, "", f"Carpeta: {paths.project_dir}", ""]
        lines += [f"  {path.relative_to(paths.project_dir).as_posix()}" for path in result.written]
        if result.backups:
            lines += ["", "Copias de los archivos reemplazados:"]
            lines += [f"  {path}" for path in result.backups]
        if any(path.name.endswith(".reports.xml") for path in result.written):
            lines += ["", "Convertido de Oracle Reports automáticamente: el SQL de cada consulta, las fórmulas,",
                      "los marcadores de posición, los totales y los Data Links que usa el Word."]
            if result.skipped:
                lines += ["", "CONSULTAS DEL MODELO DE DATOS QUE NO SE COPIARON:"]
                lines += [f"  - {note}" for note in result.skipped]
            lines += ["", "PENDIENTE A MANO:"] + ([f"  - {note}" for note in result.notes] or ["  Nada."])
            lines += [
                "",
                "Siguientes pasos:",
                "1. Resuelva los pendientes de arriba, si hay.",
                "2. Revise 'parameters' (Page Item y tipo de cada parámetro de Reports) y el título.",
                "3. Pulse Validar y luego Compilar: la validación vuelve a listar lo que falte.",
            ]
            return "\n".join(lines)
        if result.notes:
            lines += ["", "Pendiente a mano:"]
            lines += [f"  - {note}" for note in result.notes]
        lines += [
            "",
            "Siguientes pasos:",
            "1. En cada generado/q_<consulta>.sql pegue el SQL de la Query del reporte original,",
            "   con alias iguales a las columnas que usa la plantilla.",
            "2. Declare en 'parameters' cada :P_NOMBRE que usen las consultas (Page Item y tipo).",
            "3. Revise 'constants' (valores de los {{FIELD:NOMBRE}} sin consulta) y el título.",
            "4. Pulse Validar y luego Compilar.",
        ]
        return "\n".join(lines)
    assert result.skeleton is not None
    lines = ["PROYECTO PREPARADO DESDE EL DOCX", "=" * 32, "", f"Carpeta: {paths.project_dir}", ""]
    entries = (
        (paths.docx.name, "plantilla"),
        (PROCESS_NAME, "se generará al compilar; es lo que se pega en APEX"),
        (f"{GENERATED_DIRNAME}/{paths.sql.name}", "consulta: reemplace tabla_origen y columnas"),
        (f"{GENERATED_DIRNAME}/{paths.project.name}", "Page Items, título, anchos"),
    )
    width = max(len(name) for name, _ in entries)
    lines += [f"  {name.ljust(width)}  {description}" for name, description in entries]
    if result.backups:
        lines += ["", "Copias de los archivos reemplazados:"]
        lines += [f"  {path}" for path in result.backups]
    lines += ["", "Mapeo de marcadores del DOCX a Page Items:"]
    lines += _table(result.skeleton.mapping)
    lines.append("")
    if page.strip().upper() in {"", PAGE_PLACEHOLDER}:
        lines += [
            f"Los items usan el prefijo P{PAGE_PLACEHOLDER}_: reemplace {PAGE_PLACEHOLDER} por el número de",
            "página APEX en el .report.json (por ejemplo P42_).",
        ]
    lines += [
        "Siguientes pasos:",
        "1. Edite generado/<nombre>.sql: tabla real, columnas, JOIN para LOV.",
        "2. Elimine los filtros y bindings que no necesite (deben coincidir 1 a 1).",
        "3. Ajuste título, file_name y column_widths en el .report.json.",
        "4. Pulse Validar y luego Compilar.",
    ]
    return "\n".join(lines)


# ------------------------------------------------------------------ temporales
STAGING_DIR = re.compile(r"^\..+-[a-z0-9_]{8}$")     # tempfile.mkdtemp(prefix=".<nombre>-") del compilador


@dataclass(frozen=True, slots=True)
class TemporaryItem:
    path: Path
    kind: str          # descripción para el usuario
    size: int          # bytes


def _size(path: Path) -> int:
    try:
        if path.is_file():
            return path.stat().st_size
        return sum(item.stat().st_size for item in path.rglob("*") if item.is_file())
    except OSError:
        return 0


def temporary_files(projects_root: Path | None = None, *, system_temp: Path | None = None) -> list[TemporaryItem]:
    """Archivos que la herramienta deja y se pueden borrar sin perder ningún proyecto.

    - ``proyectos/_temporal/``: conversiones de rwconverter (.rdf -> .xml) que quedaron
      a medias (el XML útil ya está en ``generado/<nombre>.reports.xml``).
    - ``*.bak``: copias de los archivos que reemplazó una nueva carga del DOCX.
    - Carpetas ``.<nombre>-xxxxxxxx`` que deja una compilación interrumpida.
    - ``%TEMP%/reports-xml-*``: conversiones de versiones anteriores de la herramienta.
    """

    import tempfile

    root = Path(projects_root or default_projects_root())
    temp_dir = TEMP_DIR if projects_root is None else root / TEMP_DIR.name
    found: list[TemporaryItem] = []
    if temp_dir.is_dir():
        for path in sorted(temp_dir.rglob("*.xml")):
            found.append(TemporaryItem(path.parent if path.parent != temp_dir else path,
                                       "conversión temporal de rwconverter", _size(path)))
    if root.is_dir():
        for path in sorted(root.rglob("*")):
            if temp_dir in path.parents or path == temp_dir:
                continue
            if path.is_file() and path.suffix.lower() == ".bak":
                found.append(TemporaryItem(path, "copia .bak de una carga anterior", _size(path)))
            elif path.is_dir() and STAGING_DIR.match(path.name):
                found.append(TemporaryItem(path, "carpeta de una compilación interrumpida", _size(path)))
    legacy = Path(system_temp or tempfile.gettempdir())
    if legacy.is_dir():
        for path in sorted(legacy.glob(f"{LEGACY_TEMP_PREFIX}*")):
            if path.is_dir():
                found.append(TemporaryItem(path, "conversión de rwconverter de una versión anterior (en %TEMP%)",
                                           _size(path)))
    unique = {item.path: item for item in found}               # una carpeta con varios .xml: una sola vez
    return list(unique.values())


def remove_temporary(items: list[TemporaryItem]) -> list[str]:
    """Borra los temporales indicados; devuelve los errores (vacío si todo se borró)."""

    errors = []
    for item in items:
        try:
            if item.path.is_dir():
                shutil.rmtree(item.path)
            elif item.path.exists():
                item.path.unlink()
        except OSError as exc:
            errors.append(f"{item.path}: {exc}")
    return errors


def format_size(size: int) -> str:
    for unit in ("bytes", "KB", "MB"):
        if size < 1024 or unit == "MB":
            return f"{size} {unit}" if unit == "bytes" else f"{size:.1f} {unit}"
        size /= 1024
    return f"{size} bytes"
