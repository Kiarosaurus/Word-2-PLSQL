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
        *(paths.generated_dir / f"q_{query.lower()}.sql" for query in sorted(queries)),
    )


def files_to_replace(docx_path: Path, projects_root: Path | None = None) -> tuple[Path, ...]:
    """Archivos existentes que una carga de ``docx_path`` reemplazaría."""

    source = Path(docx_path).resolve()
    paths = WorkspacePaths.for_name((projects_root or default_projects_root()).resolve(), source.stem)
    managed = paths.managed_files()
    if is_layout_docx(source):
        template = read_layout_template(source, Diagnostics(strict=True))
        managed = _layout_managed(paths, set(template.queries) if template else set())
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


def _import_layout_docx(source: Path, root: Path, *, replace: bool) -> ImportResult:
    diagnostics = Diagnostics(strict=True)
    template = read_layout_template(source, diagnostics)
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
    identifier = _layout_identifier(source.stem)
    project = {
        "schema": "corporate-layout-project/1.0",
        "report_id": identifier,
        "template": f"../{paths.docx.name}",
        "title": re.sub(r"\s+", " ", source.stem.replace("_", " ")).strip() or "Reporte",
        "file_name": re.sub(r"[^A-Za-z0-9_-]+", "_", source.stem).strip("_") or "reporte",
        "parameters": [],
        "queries": {name: f"q_{name.lower()}.sql" for name in sorted(template.queries)},
        "constants": {name: name for name in sorted(template.fields)},
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
        for name, columns in sorted(template.queries.items()):
            query_path = paths.generated_dir / f"q_{name.lower()}.sql"
            query_path.write_text(_layout_query_skeleton(name, columns), encoding="utf-8", newline="\n")
            written.append(query_path)
        paths.project.write_text(json.dumps(project, ensure_ascii=False, indent=2) + "\n",
                                 encoding="utf-8", newline="\n")
        written.append(paths.project)
    except OSError as exc:
        diagnostics.error("IO-010", f"No se pudo preparar la carpeta del proyecto: {exc}", location=str(paths.project_dir))
        return ImportResult(diagnostics, paths=paths, backups=tuple(backups))
    return ImportResult(diagnostics, paths=paths, written=tuple(written), backups=tuple(backups), layout=True)


def import_docx(
    docx_path: str | Path,
    *,
    projects_root: str | Path | None = None,
    page: str = PAGE_PLACEHOLDER,
    replace: bool = False,
) -> ImportResult:
    """Copia el DOCX a ``proyectos/<nombre>/`` y genera el proyecto inicial.

    El DOCX se valida antes de tocar ningún archivo. Si la carpeta ya tiene
    archivos de una carga anterior, se exige ``replace=True``; entonces cada
    archivo reemplazado se mueve antes a ``generado/<archivo>.bak``.
    """

    diagnostics = Diagnostics(strict=True)
    source = Path(docx_path).expanduser().resolve()
    root = Path(projects_root or default_projects_root()).expanduser().resolve()
    if is_layout_docx(source):
        return _import_layout_docx(source, root, replace=replace)
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
