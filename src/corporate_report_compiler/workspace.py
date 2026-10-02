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
import shutil

from .apex_guide import PAGE_PLACEHOLDER, ProjectSkeleton, _table, build_skeleton
from .diagnostics import Diagnostics
from .docx_reader import read_template
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


def _same_file(first: Path, second: Path) -> bool:
    try:
        return first.exists() and second.exists() and os.path.samefile(first, second)
    except OSError:
        return False


def files_to_replace(docx_path: Path, projects_root: Path | None = None) -> tuple[Path, ...]:
    """Archivos existentes que una carga de ``docx_path`` reemplazaría."""

    source = Path(docx_path).resolve()
    paths = WorkspacePaths.for_name((projects_root or default_projects_root()).resolve(), source.stem)
    return tuple(
        path
        for path in paths.managed_files()
        if path.exists() and not (path == paths.docx and _same_file(path, source))
    )


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
    assert paths is not None and result.skeleton is not None
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
        "3. Ajuste título, file_name, column_widths y max_rows en el .report.json.",
        "4. Pulse Validar y luego Compilar.",
    ]
    return "\n".join(lines)
