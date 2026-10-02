from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
import traceback
from typing import Sequence

from . import __version__
from .compiler import CompilationResult, compile_project, validate_project


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
    compile_command.add_argument("--output", required=True, type=Path, help="Carpeta de salida.")
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
        if arguments.command == "validate":
            result = validate_project(arguments.project, strict=arguments.strict)
        else:
            result = compile_project(
                arguments.project,
                arguments.output,
                strict=arguments.strict,
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
