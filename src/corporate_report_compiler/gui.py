from __future__ import annotations

import os
from pathlib import Path
import re
import sys
import traceback

from .apex_guide import PAGE_PLACEHOLDER, build_apex_guide, format_skeleton, write_skeleton
from .compiler import CompilationResult, compile_project, validate_project


def _format_diagnostics(items: list[dict]) -> list[str]:
    lines: list[str] = []
    for item in items:
        location = f" — {item['location']}" if "location" in item else ""
        lines.append(f"{item['severity']} {item['code']}{location}")
        lines.append(str(item["message"]))
        if "suggestion" in item:
            lines.append(f"Solución: {item['suggestion']}")
        lines.append("")
    return lines


def _format_result(result: CompilationResult, *, compiled: bool = False) -> str:
    lines = _format_diagnostics(result.diagnostics.to_dict()["diagnostics"])
    if result.valid:
        if compiled:
            lines.append("Compilación correcta. Archivos generados:")
            lines.extend(f"  {path}" for path in result.artifacts)
            lines.append("")
            lines.append(build_apex_guide(result))
        else:
            lines.append("Validación correcta; no se escribió ningún archivo.")
            lines.append("Pulse Compilar para generar apex_process.sql y ver qué subir a APEX.")
    elif not lines:
        lines.append("La validación falló sin diagnóstico disponible.")
    else:
        lines.append(
            f"No válido: {result.diagnostics.error_count} error(es), "
            f"{result.diagnostics.warning_count} advertencia(s)."
        )
    return "\n".join(lines)


def default_output(project_path: Path) -> Path:
    """Carpeta ``build/<reporte>`` junto al proyecto, sin el sufijo ``.report``."""

    name = project_path.name
    for suffix in (".report.json", ".json"):
        if name.lower().endswith(suffix):
            name = name[: -len(suffix)]
            break
    return project_path.parent / "build" / (name or "reporte")


class CompilerApplication:
    """GUI mínima; toda la lógica permanece en :mod:`compiler` y :mod:`apex_guide`."""

    def __init__(self, root: object) -> None:
        import tkinter as tk
        from tkinter import ttk

        self.root = root
        self.tk = tk
        self.ttk = ttk
        root.title("Compilador de reportes APEX")
        root.minsize(900, 600)

        self.project = tk.StringVar()
        self.output = tk.StringVar()
        self.strict = tk.BooleanVar(value=True)
        self._output_chosen_for: str | None = None
        self._last_process: Path | None = None
        root.report_callback_exception = self._report_exception

        frame = ttk.Frame(root, padding=14)
        frame.grid(row=0, column=0, sticky="nsew")
        root.rowconfigure(0, weight=1)
        root.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(4, weight=1)

        ttk.Label(frame, text="Proyecto (.report.json)").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.project).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(frame, text="Examinar…", command=self._choose_project).grid(row=0, column=2)

        ttk.Label(frame, text="Carpeta de salida").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.output).grid(row=1, column=1, sticky="ew", padx=8)
        ttk.Button(frame, text="Examinar…", command=self._choose_output).grid(row=1, column=2)

        ttk.Checkbutton(
            frame,
            text="Modo estricto (recomendado)",
            variable=self.strict,
        ).grid(row=2, column=1, sticky="w", pady=6)

        actions = ttk.Frame(frame)
        actions.grid(row=3, column=0, columnspan=3, sticky="w", pady=(4, 10))
        ttk.Button(actions, text="Nuevo proyecto desde DOCX…", command=self._new_from_docx).pack(side="left", padx=(0, 16))
        ttk.Button(actions, text="Validar", command=self._validate).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Compilar", command=self._compile).pack(side="left", padx=(0, 16))
        self.copy_button = ttk.Button(actions, text="Copiar código APEX", command=self._copy_process, state="disabled")
        self.copy_button.pack(side="left", padx=(0, 8))
        self.open_button = ttk.Button(actions, text="Abrir carpeta de salida", command=self._open_output, state="disabled")
        self.open_button.pack(side="left")

        text_frame = ttk.Frame(frame)
        text_frame.grid(row=4, column=0, columnspan=3, sticky="nsew")
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)
        self.result = tk.Text(text_frame, wrap="none", state="disabled", font=("Consolas", 9))
        self.result.grid(row=0, column=0, sticky="nsew")
        vertical = ttk.Scrollbar(text_frame, orient="vertical", command=self.result.yview)
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal = ttk.Scrollbar(text_frame, orient="horizontal", command=self.result.xview)
        horizontal.grid(row=1, column=0, sticky="ew")
        self.result.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)

        self._show(
            "1. Si aún no tiene proyecto, pulse «Nuevo proyecto desde DOCX…».\n"
            "2. Seleccione el .report.json, pulse Validar y luego Compilar.\n"
            "3. Tras compilar, aquí verá exactamente qué subir a APEX y dónde."
        )

    # -- utilidades -------------------------------------------------------
    def _output_is_custom(self) -> bool:
        """La carpeta elegida a mano solo vale para el proyecto con que se eligió."""

        return self._output_chosen_for is not None and self._output_chosen_for == self.project.get().strip()

    def _set_project(self, path: Path) -> None:
        self.project.set(str(path))
        if not self._output_is_custom():
            self.output.set(str(default_output(path)))

    def _report_exception(self, exc_type, exc, tb) -> None:  # pragma: no cover - barrera de la GUI
        self._show(
            f"Error interno inesperado: {exc}\n\n" + "".join(traceback.format_exception(exc_type, exc, tb))
        )

    def _show(self, value: str) -> None:
        self.result.configure(state="normal")
        self.result.delete("1.0", "end")
        self.result.insert("1.0", value)
        self.result.configure(state="disabled")
        self.root.update_idletasks()

    def _run(self, action) -> None:
        """Ejecuta una acción mostrando cualquier fallo, nunca un resultado anterior."""

        self._show("Procesando…")
        self.copy_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        try:
            action()
        except Exception as exc:  # pragma: no cover - barrera de la GUI
            self._show(
                f"Error interno inesperado: {exc}\n\n"
                "No se modificó la última salida válida.\n\n" + traceback.format_exc()
            )

    def _path_or_message(self) -> Path | None:
        value = self.project.get().strip()
        if not value:
            self._show("Seleccione primero un archivo .report.json.")
            return None
        return Path(value)

    # -- acciones ---------------------------------------------------------
    def _choose_project(self) -> None:
        from tkinter import filedialog

        value = filedialog.askopenfilename(
            title="Seleccione el proyecto",
            filetypes=(("Proyecto de reporte", "*.report.json"), ("JSON", "*.json"), ("Todos", "*.*")),
        )
        if value:
            self._set_project(Path(value))

    def _choose_output(self) -> None:
        from tkinter import filedialog

        value = filedialog.askdirectory(title="Seleccione la carpeta de salida")
        if value:
            self._output_chosen_for = self.project.get().strip()
            self.output.set(value)

    def _new_from_docx(self) -> None:
        from tkinter import filedialog, simpledialog

        value = filedialog.askopenfilename(
            title="Seleccione la plantilla Word",
            filetypes=(("Documento de Word", "*.docx"), ("Todos", "*.*")),
        )
        if not value:
            return
        page = simpledialog.askstring(
            "Número de página APEX",
            "Número de la página APEX que tendrá los filtros.\n"
            f"Deje {PAGE_PLACEHOLDER} para generar items P{PAGE_PLACEHOLDER}_ y reemplazarlo después.",
            initialvalue=PAGE_PLACEHOLDER,
            parent=self.root,
        )
        if page is None:
            return
        page = page.strip().upper() or PAGE_PLACEHOLDER
        if page != PAGE_PLACEHOLDER and not re.fullmatch(r"0|[1-9][0-9]{0,5}", page):
            self._show("El número de página debe ser numérico (por ejemplo 42) o XX.")
            return

        def action() -> None:
            diagnostics, skeleton, files = write_skeleton(Path(value), page=page)
            lines = _format_diagnostics(diagnostics.to_dict()["diagnostics"])
            if skeleton is not None and files:
                self._set_project(files[0])
                lines.append(format_skeleton(skeleton, files, page))
            elif not lines:
                lines.append("No se pudo generar el proyecto.")
            self._show("\n".join(lines))

        self._run(action)

    def _validate(self) -> None:
        path = self._path_or_message()
        if path is not None:
            self._run(lambda: self._show(_format_result(validate_project(path, strict=self.strict.get()))))

    def _compile(self) -> None:
        path = self._path_or_message()
        if path is None:
            return
        output = self.output.get().strip()
        if not output or not self._output_is_custom():
            output = str(default_output(path))
            self.output.set(output)

        def action() -> None:
            result = compile_project(path, Path(output), strict=self.strict.get())
            self._show(_format_result(result, compiled=True))
            self._last_process = next((item for item in result.artifacts if item.name == "apex_process.sql"), None)
            if self._last_process is not None:
                self.copy_button.configure(state="normal")
                self.open_button.configure(state="normal")

        self._run(action)

    def _copy_process(self) -> None:
        if self._last_process is None or not self._last_process.is_file():
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self._last_process.read_text(encoding="utf-8"))

    def _open_output(self) -> None:
        if self._last_process is None:
            return
        folder = self._last_process.parent
        if sys.platform.startswith("win"):
            os.startfile(folder)  # type: ignore[attr-defined]  # noqa: S606


def _fatal(message: str) -> None:
    """Informa un fallo de arranque también cuando no hay consola (pythonw)."""

    if sys.stderr is not None:
        print(message, file=sys.stderr)
    if sys.platform.startswith("win"):
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(None, message, "Compilador de reportes APEX", 0x10)
        except (AttributeError, OSError):
            pass


def main() -> int:
    try:
        import tkinter as tk
    except ImportError as exc:
        _fatal(f"No se pudo cargar Tkinter: {exc}\nReinstale Python incluyendo «tcl/tk and IDLE».")
        return 3
    try:
        root = tk.Tk()
        CompilerApplication(root)
        root.mainloop()
        return 0
    except tk.TclError as exc:
        _fatal(f"No se pudo iniciar la interfaz gráfica: {exc}")
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
