from __future__ import annotations

from pathlib import Path
import sys

from .compiler import CompilationResult, compile_project, validate_project


def _format_result(result: CompilationResult) -> str:
    lines: list[str] = []
    for item in result.diagnostics.to_dict()["diagnostics"]:
        location = f" — {item['location']}" if "location" in item else ""
        lines.append(f"{item['severity']} {item['code']}{location}")
        lines.append(str(item["message"]))
        if "suggestion" in item:
            lines.append(f"Solución: {item['suggestion']}")
        lines.append("")
    if result.valid:
        lines.append("Resultado válido.")
        lines.extend(f"Generado: {path}" for path in result.artifacts)
    elif not lines:
        lines.append("La validación falló sin diagnóstico disponible.")
    return "\n".join(lines)


class CompilerApplication:
    """GUI mínima; toda la lógica permanece en :mod:`compiler`."""

    def __init__(self, root: object) -> None:
        import tkinter as tk
        from tkinter import ttk

        self.root = root
        self.tk = tk
        self.ttk = ttk
        root.title("Compilador de reportes APEX")
        root.minsize(760, 520)

        self.project = tk.StringVar()
        self.output = tk.StringVar(value=str(Path.cwd() / "compiled"))
        self.strict = tk.BooleanVar(value=True)

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
        ttk.Button(actions, text="Validar", command=self._validate).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Compilar", command=self._compile).pack(side="left")

        self.result = tk.Text(frame, wrap="word", state="disabled")
        self.result.grid(row=4, column=0, columnspan=3, sticky="nsew")

    def _choose_project(self) -> None:
        from tkinter import filedialog

        value = filedialog.askopenfilename(
            title="Seleccione el proyecto",
            filetypes=(("Proyecto de reporte", "*.report.json"), ("JSON", "*.json"), ("Todos", "*.*")),
        )
        if value:
            self.project.set(value)
            if self.output.get().endswith("compiled"):
                self.output.set(str(Path(value).parent / "compiled" / Path(value).stem))

    def _choose_output(self) -> None:
        from tkinter import filedialog

        value = filedialog.askdirectory(title="Seleccione la carpeta de salida")
        if value:
            self.output.set(value)

    def _show(self, value: str) -> None:
        self.result.configure(state="normal")
        self.result.delete("1.0", "end")
        self.result.insert("1.0", value)
        self.result.configure(state="disabled")

    def _path_or_message(self) -> Path | None:
        value = self.project.get().strip()
        if not value:
            self._show("Seleccione primero un archivo .report.json.")
            return None
        return Path(value)

    def _validate(self) -> None:
        path = self._path_or_message()
        if path is not None:
            self._show(_format_result(validate_project(path, strict=self.strict.get())))

    def _compile(self) -> None:
        path = self._path_or_message()
        output = self.output.get().strip()
        if path is None:
            return
        if not output:
            self._show("Seleccione una carpeta de salida.")
            return
        self._show(
            _format_result(
                compile_project(path, Path(output), strict=self.strict.get())
            )
        )


def main() -> int:
    try:
        import tkinter as tk
    except ImportError as exc:
        print(f"No se pudo cargar Tkinter: {exc}", file=sys.stderr)
        return 3
    try:
        root = tk.Tk()
        CompilerApplication(root)
        root.mainloop()
        return 0
    except tk.TclError as exc:
        print(f"No se pudo iniciar la interfaz gráfica: {exc}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
