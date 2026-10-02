from __future__ import annotations

import argparse
import json
from pathlib import Path
import re
import sys
import traceback
from typing import Sequence

from . import __version__
from .apex_guide import PAGE_PLACEHOLDER
from .compiler import CompilationResult, compile_project, validate_project
from .workspace import format_import, import_docx, workspace_outputs


def _add_validation_options(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--project", required=True, type=Path, help="Archivo .report.json del proyecto.")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument(
        "--strict",
        dest="strict",
        action="store_true",
        default=True,
        help="Rechaza formato no representable (predeterminado).",
    )
    mode.add_argument(
        "--compatible",
        dest="strict",
        action="store_false",
        help="Acepta aproximaciones documentadas como advertencias.",
    )
    parser.add_argument("--json", action="store_true", help="Imprime diagnósticos como JSON.")
    parser.add_argument("--debug", action="store_true", help="Muestra traceback ante un fallo interno.")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="apex-report-compiler",
        description="Compila plantillas Word restringidas para reportes Oracle APEX.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    commands = parser.add_subparsers(dest="command", required=True)

    validate = commands.add_parser("validate", help="Valida sin escribir artefactos.")
    _add_validation_options(validate)

    compile_command = commands.add_parser("compile", help="Valida y genera los artefactos APEX.")
    _add_validation_options(compile_command)
    compile_command.add_argument(
        "--output",
        type=Path,
        help=(
            "Carpeta de salida. Opcional si el proyecto está en proyectos/<nombre>/generado/: "
            "entonces apex_process.sql va a proyectos/<nombre>/ y el resto a generado/."
        ),
    )

    new = commands.add_parser("new", help="Copia un DOCX a proyectos/<nombre>/ y crea su proyecto inicial.")
    new.add_argument("--docx", required=True, type=Path, help="Plantilla Word.")
    new.add_argument("--page", default=PAGE_PLACEHOLDER, help="Número de página APEX de los filtros (por defecto XX).")
    new.add_argument("--projects-dir", type=Path, help="Carpeta de proyectos (por defecto proyectos/ del compilador).")
    new.add_argument(
        "--replace",
        action="store_true",
        help="Reemplaza los archivos de una carga anterior; cada uno se guarda como generado/<archivo>.bak.",
    )
    new.add_argument("--json", action="store_true", help="Imprime diagnósticos como JSON.")
    new.add_argument("--debug", action="store_true", help="Muestra traceback ante un fallo interno.")
    return parser


def _print_human(result: CompilationResult, *, command: str) -> None:
    for item in result.diagnostics.to_dict()["diagnostics"]:
        location = f" — {item['location']}" if "location" in item else ""
        print(
            f"{item['severity']} {item['code']}{location}\n{item['message']}",
            file=sys.stderr,
        )
        if "suggestion" in item:
            print(f"Solución: {item['suggestion']}", file=sys.stderr)

    if result.valid:
        if command == "compile":
            print("Compilación correcta.")
            for artifact in result.artifacts:
                print(f"  {artifact}")
        else:
            print("Validación correcta; no se escribió ningún artefacto.")
    else:
        print(
            f"No válido: {result.diagnostics.error_count} error(es), "
            f"{result.diagnostics.warning_count} advertencia(s).",
            file=sys.stderr,
        )


def _print_json(result: CompilationResult) -> None:
    payload = result.diagnostics.to_dict()
    payload["artifacts"] = [str(path) for path in result.artifacts]
    print(json.dumps(payload, ensure_ascii=False, indent=2))


def _exit_code(result: CompilationResult) -> int:
    if result.valid:
        return 0
    if any(item.code.startswith(("IO-", "EMIT-")) for item in result.diagnostics.items):
        return 3
    return 1


def _new_project(arguments: argparse.Namespace) -> int:
    page = str(arguments.page).strip().upper() or PAGE_PLACEHOLDER
    if page != PAGE_PLACEHOLDER and not re.fullmatch(r"0|[1-9][0-9]{0,5}", page):
        print("El número de página debe ser numérico (por ejemplo 42) o XX.", file=sys.stderr)
        return 2
    result = import_docx(
        arguments.docx,
        projects_root=arguments.projects_dir,
        page=page,
        replace=arguments.replace,
    )
    if arguments.json:
        payload = result.diagnostics.to_dict()
        payload["files"] = [str(path) for path in result.written]
        payload["backups"] = [str(path) for path in result.backups]
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        for item in result.diagnostics.to_dict()["diagnostics"]:
            location = f" — {item['location']}" if "location" in item else ""
            print(f"{item['severity']} {item['code']}{location}\n{item['message']}", file=sys.stderr)
        if result.written:
            print(format_import(result, page))
    if not result.written:
        return 3 if any(item.code.startswith("IO-") for item in result.diagnostics.items) else 1
    return 0


def _configure_streams() -> None:
    """Evita fallos de codificación al redirigir la salida en Windows.

    Con una tubería, Python usa la página ANSI (cp1252) y un carácter fuera de
    ella interrumpía la CLI después de escribir los artefactos.
    """

    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            try:
                reconfigure(encoding="utf-8", errors="backslashreplace")
            except (OSError, ValueError):
                pass


def main(argv: Sequence[str] | None = None) -> int:
    _configure_streams()
    parser = build_parser()
    arguments = parser.parse_args(argv)
    try:
        if arguments.command == "new":
            return _new_project(arguments)
        if arguments.command == "validate":
            result = validate_project(arguments.project, strict=arguments.strict)
        else:
            output, process_directory = arguments.output, None
            if output is None:
                layout = workspace_outputs(arguments.project.resolve())
                if layout is None:
                    parser.error("--output es obligatorio fuera de proyectos/<nombre>/generado/.")
                output, process_directory = layout
            result = compile_project(
                arguments.project,
                output,
                strict=arguments.strict,
                process_directory=process_directory,
            )
    except Exception as exc:  # pragma: no cover - última barrera de la CLI
        if arguments.debug:
            traceback.print_exc()
        else:
            print(f"Error interno inesperado: {exc}", file=sys.stderr)
        return 4

    if arguments.json:
        _print_json(result)
    else:
        _print_human(result, command=arguments.command)
    return _exit_code(result)


if __name__ == "__main__":
    raise SystemExit(main())
