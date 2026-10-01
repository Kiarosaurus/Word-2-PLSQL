from __future__ import annotations

from dataclasses import dataclass, replace
import json
import os
from pathlib import Path
import shutil
import tempfile
from typing import Any

from .diagnostics import Diagnostics
from .docx_reader import read_template
from .emitter import build_definition, emit_artifacts
from .model import ProjectModel, TemplateModel
from .project import load_project, locate_project_inputs, read_sql


@dataclass(frozen=True, slots=True)
class CompilationResult:
    """Resultado estable compartido por la API, la CLI y la GUI.

    ``validate_project`` nunca escribe archivos. ``compile_project`` solo
    completa ``artifacts`` después de que los tres artefactos hayan sido
    generados, verificados y trasladados al destino.
    """

    project_path: Path
    diagnostics: Diagnostics
    template: TemplateModel | None = None
    project: ProjectModel | None = None
    sql: str | None = None
    definition: dict[str, Any] | None = None
    artifacts: tuple[Path, ...] = ()

    @property
    def valid(self) -> bool:
        return not self.diagnostics.has_errors

    @property
    def success(self) -> bool:
        """Alias cómodo para consumidores que tratan la compilación como tarea."""

        return self.valid

    @property
    def output_files(self) -> tuple[Path, ...]:
        """Alias compatible y explícito de :attr:`artifacts`."""

        return self.artifacts


def validate_project(
    project_path: str | Path,
    *,
    strict: bool = True,
) -> CompilationResult:
    """Valida un proyecto completo sin crear ni modificar archivos de salida."""

    path = Path(project_path).expanduser().resolve()
    diagnostics = Diagnostics(strict=strict)
    _raw, template_path, query_path = locate_project_inputs(path, diagnostics)

    template: TemplateModel | None = None
    sql: str | None = None
    sql_binds: tuple[str, ...] = ()

    # DOCX y SQL son independientes: leer ambos permite devolver en una sola
    # ejecución todos los diagnósticos que puedan obtenerse con seguridad.
    if template_path is not None:
        template = read_template(template_path, diagnostics)
    if query_path is not None:
        sql, sql_binds = read_sql(query_path, diagnostics)

    project: ProjectModel | None = None
    definition: dict[str, Any] | None = None
    if template is not None and sql is not None:
        project = load_project(path, template, sql, sql_binds, diagnostics)
        if project is not None and not diagnostics.has_errors:
            definition = build_definition(template, project, sql)

    if diagnostics.has_errors:
        project = None
        definition = None

    return CompilationResult(
        project_path=path,
        diagnostics=diagnostics,
        template=template,
        project=project,
        sql=sql,
        definition=definition,
    )


def _verify_staged_artifacts(
    staging: Path,
    definition: dict[str, Any],
    diagnostics: Diagnostics,
) -> tuple[Path, ...]:
    expected = (
        staging / "template.json",
        staging / "validation.json",
        staging / "apex_process.sql",
    )
    for path in expected:
        if not path.is_file() or path.stat().st_size == 0:
            diagnostics.error(
                "EMIT-001",
                f"No se generó correctamente {path.name}.",
                location=str(path),
            )

    if diagnostics.has_errors:
        return ()

    try:
        emitted_definition = json.loads(expected[0].read_text(encoding="utf-8"))
        emitted_validation = json.loads(expected[1].read_text(encoding="utf-8"))
        emitted_process = expected[2].read_text(encoding="utf-8")
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        diagnostics.error("EMIT-002", f"No se pudo verificar la salida: {exc}")
        return ()

    if emitted_definition != definition:
        diagnostics.error("EMIT-003", "template.json no coincide con el modelo compilado.")
    if emitted_validation.get("valid") is not True:
        diagnostics.error("EMIT-004", "validation.json no confirmó una compilación válida.")
    if "pkg_corporate_reports.download_query" not in emitted_process.lower():
        diagnostics.error("EMIT-005", "apex_process.sql no llama a DOWNLOAD_QUERY.")
    return expected if not diagnostics.has_errors else ()


def _commit_artifacts(staging_files: tuple[Path, ...], output: Path) -> tuple[Path, ...]:
    """Publica los archivos con rollback si falla una sustitución.

    El directorio destino puede contener otros archivos: solo se sustituyen los
    tres nombres contractuales. Los archivos anteriores se conservan en un
    respaldo temporal hasta completar toda la operación.
    """

    output.mkdir(parents=True, exist_ok=True)
    backup = Path(tempfile.mkdtemp(prefix=".report-compiler-backup-", dir=output.parent))
    installed: list[Path] = []
    saved: list[tuple[Path, Path]] = []
    try:
        for source in staging_files:
            destination = output / source.name
            if destination.exists():
                backup_path = backup / source.name
                os.replace(destination, backup_path)
                saved.append((backup_path, destination))
            os.replace(source, destination)
            installed.append(destination)
    except Exception:
        for destination in reversed(installed):
            try:
                destination.unlink(missing_ok=True)
            except OSError:
                pass
        for backup_path, destination in reversed(saved):
            if backup_path.exists():
                os.replace(backup_path, destination)
        raise
    finally:
        shutil.rmtree(backup, ignore_errors=True)
    return tuple(output / path.name for path in staging_files)


def compile_project(
    project_path: str | Path,
    output_directory: str | Path,
    *,
    strict: bool = True,
) -> CompilationResult:
    """Valida y compila atómicamente un proyecto.

    Una validación fallida no crea el directorio de salida ni altera una
    compilación anterior. Los errores de E/S se convierten en diagnósticos
    controlados; la CLI decide el código de salida correspondiente.
    """

    result = validate_project(project_path, strict=strict)
    if not result.valid:
        return result
    assert result.definition is not None
    assert result.project is not None

    output = Path(output_directory).expanduser().resolve()
    if output.exists() and not output.is_dir():
        result.diagnostics.error(
            "IO-001",
            "La ruta de salida existe y no es una carpeta.",
            location=str(output),
        )
        return replace(result, definition=None)

    output_parent = output.parent
    staging: Path | None = None
    try:
        output_parent.mkdir(parents=True, exist_ok=True)
        staging = Path(tempfile.mkdtemp(prefix=f".{output.name or 'compiled'}-", dir=output_parent))
        emit_artifacts(
            staging,
            result.definition,
            result.project,
            result.diagnostics,
        )
        staged_files = _verify_staged_artifacts(staging, result.definition, result.diagnostics)
        if result.diagnostics.has_errors:
            return replace(result, definition=None)
        artifacts = _commit_artifacts(staged_files, output)
        return replace(result, artifacts=artifacts)
    except OSError as exc:
        result.diagnostics.error(
            "IO-002",
            f"No se pudieron escribir los artefactos: {exc}",
            location=str(output),
        )
        return replace(result, definition=None, artifacts=())
    finally:
        if staging is not None:
            shutil.rmtree(staging, ignore_errors=True)
