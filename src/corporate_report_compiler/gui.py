from __future__ import annotations

import os
from pathlib import Path
import re
import sys
import traceback

from .apex_guide import PAGE_PLACEHOLDER, build_apex_guide, build_detailed_instructions, process_code
from .compiler import CompilationResult, compile_project, validate_project
from .workspace import (
    default_projects_root,
    existing_page,
    files_to_replace,
    format_import,
    import_docx,
    format_size,
    page_placeholder_locations,
    remove_temporary,
    temporary_files,
    workspace_outputs,
)
from .reports_model import TEMP_DIR, discard_conversion


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


def _placeholder_locations(project_path: Path | None) -> list[str]:
    """Dónde reemplazar P{XX}_ por el número de página: archivo, línea e items."""

    if project_path is None:
        return []
    layout = workspace_outputs(Path(project_path))
    folder = layout[1] if layout is not None else Path(project_path).parent
    found = page_placeholder_locations(folder)
    if not found:
        return []
    lines = [f"Reemplace {PAGE_PLACEHOLDER} por el número de página en:"]
    for path, number, items in found:
        try:
            shown = path.relative_to(folder).as_posix()
        except ValueError:
            shown = str(path)
        lines.append(f"  {shown}:{number}  " + ", ".join(items))
    return lines + ["  (apex_process.sql y layout.json/template.json se rehacen desde el .report.json al Compilar.)"]


def _page_items_summary(project_path: Path | None) -> list[str]:
    """Recordatorio de la página APEX elegida y de los Page Items que el reporte usa."""

    import json

    try:
        raw = json.loads(Path(project_path).read_text(encoding="utf-8-sig"))
    except (OSError, TypeError, ValueError):
        return []
    if not isinstance(raw, dict):
        return []
    items: list[tuple[str, str]] = []                 # (item, descripción)
    for entry in raw.get("parameters") or raw.get("bindings") or []:
        if isinstance(entry, dict) and entry.get("item"):
            name = entry.get("name") or entry.get("bind") or ""
            kind = str(entry.get("type") or "VARCHAR2").upper()
            kind += " obligatorio" if entry.get("required") else ""
            items.append((str(entry["item"]).upper(), f"{kind}  -> :{str(name).upper()}"))
    for key, label in (("format_item", "formato PDF/XLSX"), ("orientation_item", "orientación")):
        if raw.get(key):
            items.append((str(raw[key]).upper(), label))
    pages = sorted({match.group(1) for item, _ in items if (match := re.match(r"P(\d+|XX)_", item))},
                   key=lambda page: (page == "0", page))
    main = [page for page in pages if page != "0"]
    lines = ["PÁGINA APEX Y PAGE ITEMS", "=" * 24]
    if not main:
        lines.append("Página elegida: ninguna (el reporte no tiene filtros propios de una página).")
    elif main == [PAGE_PLACEHOLDER]:
        lines.append(f"Página elegida: {PAGE_PLACEHOLDER} (sin número todavía: reemplace {PAGE_PLACEHOLDER} "
                     "por la página real o vuelva a cargar el Word indicando la página).")
    else:
        lines.append("Página elegida: " + ", ".join(main)
                     + (" (¡items de varias páginas!)" if len(main) > 1 else ""))
    # También fuera del .report.json: un q_*.sql o una nota editada a mano puede seguir con P{XX}_.
    lines += _placeholder_locations(project_path)
    if items:
        width = max(len(item) for item, _ in items)
        lines.append("Page Items que deben existir en APEX:")
        kind_width = max(len(description.split("  -> ")[0]) for _, description in items)
        lines += [f"  {item:<{width}}  " + ("  -> ".join(
            [description.split("  -> ")[0].ljust(kind_width)] + description.split("  -> ")[1:])).rstrip()
            for item, description in items]
    else:
        lines.append("Page Items: ninguno.")
    return lines + ["", ""]


def _reports_summary(result: CompilationResult) -> list[str]:
    """Resumen de lo convertido de Oracle Reports y de lo que falta insertar a mano."""

    pending = [item for item in result.diagnostics.to_dict()["diagnostics"]
               if item["code"].startswith("LAYOUT-REPORTS-")]
    reports = (result.definition or {}).get("reports")
    if not pending and not reports:
        return []
    lines = ["CÓDIGO DE ORACLE REPORTS", "=" * 24]
    if reports:
        model = reports["model"]
        lines.append(
            f"Convertido automáticamente: {len(model['formulas'])} fórmula(s), "
            f"{len(model['placeholders'])} marcador(es) de posición, {len(model['summaries'])} total(es) y "
            f"{sum(len(q.get('filters', [])) for q in result.definition['queries'].values())} Data Link(s)."
        )
    if pending:
        lines.append("FALTA INSERTAR O REVISAR:")
        lines += [f"  - [{item['code']}] {item['message']}" for item in pending]
    else:
        lines.append("Nada pendiente a mano.")
    return lines + ["", ""]


def _format_result(result: CompilationResult, *, compiled: bool = False) -> str:
    lines = (_page_items_summary(result.project_path) + _reports_summary(result)
             + _format_diagnostics(result.diagnostics.to_dict()["diagnostics"]))
    if result.valid:
        if compiled:
            lines.append("Compilación correcta. Archivos generados:")
            lines.extend(f"  {path}" for path in result.artifacts)
            lines.append("")
            lines.append(build_apex_guide(result))
        else:
            lines.append("Validación correcta; no se escribió ningún archivo.")
            lines.append("Pulse Compilar para generar o actualizar los archivos citados abajo. «Copiar código APEX»")
            lines.append("ya copia el proceso; «Instrucciones APEX detalladas» explica cada paso.")
            lines.append("")
            lines.append(build_apex_guide(result))
    elif not lines:
        lines.append("La validación falló sin diagnóstico disponible.")
    else:
        lines.append(
            f"No válido: {result.diagnostics.error_count} error(es), "
            f"{result.diagnostics.warning_count} advertencia(s)."
        )
        lines.append("Al corregirlos, Validar y Compilar muestran qué subir a APEX y dónde; "
                     "vea también «Instrucciones APEX detalladas».")
    return "\n".join(lines)


def default_output(project_path: Path) -> Path:
    """Carpeta de salida predeterminada.

    En ``proyectos/<nombre>/generado/`` es esa misma carpeta (y
    ``apex_process.sql`` va a ``proyectos/<nombre>/``). Fuera de esa
    organización, ``build/<reporte>`` junto al proyecto.
    """

    layout = workspace_outputs(project_path)
    if layout is not None:
        return layout[0]
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
        self._process_text: str | None = None      # proceso generado en memoria (tras Validar)
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
        self.open_button.pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Instrucciones APEX detalladas", command=self._instructions).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Archivos temporales…", command=self._temporary).pack(side="left")

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
            "1. Si aún no tiene proyecto, pulse «Nuevo proyecto desde DOCX…»: el Word se copia a\n"
            f"   {default_projects_root()}\\<nombre>\\ y el proyecto se crea en su carpeta generado.\n"
            "2. Seleccione el .report.json (carpeta generado), pulse Validar y luego Compilar.\n"
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
            initialdir=str(default_projects_root()),
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
        previous = existing_page(Path(value))
        page = simpledialog.askstring(
            "Número de página APEX",
            "Número de la página APEX que tendrá los filtros.\n"
            f"Deje {PAGE_PLACEHOLDER} para generar items P{PAGE_PLACEHOLDER}_ y reemplazarlo después."
            + (f"\n\nEste DOCX ya es un proyecto con la página {previous}: se propone la misma "
               f"para no volver a P{PAGE_PLACEHOLDER}_." if previous else ""),
            initialvalue=previous or PAGE_PLACEHOLDER,
            parent=self.root,
        )
        if page is None:
            return
        page = page.strip().upper() or PAGE_PLACEHOLDER
        if page != PAGE_PLACEHOLDER and not re.fullmatch(r"0|[1-9][0-9]{0,5}", page):
            self._show("El número de página debe ser numérico (por ejemplo 42) o XX.")
            return

        from tkinter import messagebox

        reports = None
        if messagebox.askyesno(
            "Oracle Reports original",
            "¿Tiene el .rdf (o el .xml) del reporte original de Oracle Reports?\n\n"
            "Si lo indica, en el Word basta escribir el nombre del campo o de su Origen "
            "({{FIELD:F_96}}, {{COLUMN:INTERES2}}, {{SUM:F_108}}): la herramienta averigua la "
            "consulta de cada dato y copia el SQL real de cada consulta.",
            parent=self.root,
        ):
            chosen = filedialog.askopenfilename(
                title="Seleccione el reporte de Oracle Reports",
                filetypes=(("Oracle Reports", "*.rdf *.RDF *.xml"), ("Todos", "*.*")),
            )
            if not chosen:
                return
            reports = Path(chosen)
            self._show("Leyendo el Modelo de Datos de Oracle Reports…")

        existing = files_to_replace(Path(value), reports=reports)
        if existing:
            listing = "\n".join(f"  {path.name}" for path in existing)
            if not messagebox.askyesno(
                "Reemplazar proyecto",
                f"La carpeta del proyecto «{Path(value).stem}» ya tiene estos archivos:\n{listing}\n\n"
                "Se reemplazarán y de cada uno se guardará una copia .bak en la carpeta generado. "
                "Los demás archivos de la carpeta no se tocan.\n\n¿Continuar?",
                parent=self.root,
            ):
                discard_conversion(reports)
                return

        def action() -> None:
            result = import_docx(Path(value), page=page, replace=bool(existing), reports=reports)
            lines = _format_diagnostics(result.diagnostics.to_dict()["diagnostics"])
            if result.written and result.paths is not None:
                self._set_project(result.paths.project)
                lines = _page_items_summary(result.paths.project) + lines
                lines.append(format_import(result, page))
            elif not lines:
                lines.append("No se pudo generar el proyecto.")
            self._show("\n".join(lines))

        self._run(action)

    def _remember_process(self, result: CompilationResult) -> None:
        """Deja listo «Copiar código APEX» (y la carpeta) tanto al validar como al compilar."""

        from .apex_guide import expected_artifact

        self._process_text = process_code(result)
        self._last_process = expected_artifact(result, "apex_process.sql")
        self.copy_button.configure(state="normal" if self._process_text else "disabled")
        folder_exists = self._last_process is not None and self._last_process.parent.is_dir()
        self.open_button.configure(state="normal" if folder_exists else "disabled")

    def _validate(self) -> None:
        path = self._path_or_message()
        if path is None:
            return

        def action() -> None:
            result = validate_project(path, strict=self.strict.get())
            self._show(_format_result(result))
            self._remember_process(result)

        self._run(action)

    def _compile(self) -> None:
        path = self._path_or_message()
        if path is None:
            return
        output = self.output.get().strip()
        if not output or not self._output_is_custom():
            output = str(default_output(path))
            self.output.set(output)

        def action() -> None:
            layout = workspace_outputs(path)
            process_directory = layout[1] if layout is not None and not self._output_is_custom() else None
            result = compile_project(
                path,
                Path(output),
                strict=self.strict.get(),
                process_directory=process_directory,
            )
            self._show(_format_result(result, compiled=True))
            self._remember_process(result)

        self._run(action)

    def _copy_process(self) -> None:
        if not self._process_text:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self._process_text)

    def _instructions(self) -> None:
        """Ventana con los pasos detallados de APEX para el proyecto actual (simple o layout)."""

        path = self._path_or_message()
        if path is None:
            return
        result = validate_project(path, strict=self.strict.get())
        self._remember_process(result)
        text = build_detailed_instructions(result)
        window = self.tk.Toplevel(self.root)
        window.title("Instrucciones APEX detalladas")
        window.minsize(860, 560)
        bar = self.ttk.Frame(window, padding=(10, 8))
        bar.pack(side="top", fill="x")

        def copy(value: str) -> None:
            window.clipboard_clear()
            window.clipboard_append(value)

        self.ttk.Button(bar, text="Copiar código APEX", command=lambda: copy(self._process_text or ""),
                        state="normal" if self._process_text else "disabled").pack(side="left", padx=(0, 8))
        self.ttk.Button(bar, text="Copiar instrucciones", command=lambda: copy(text)).pack(side="left")
        body = self.tk.Text(window, wrap="none", font=("Consolas", 10))
        scroll_y = self.ttk.Scrollbar(window, orient="vertical", command=body.yview)
        scroll_x = self.ttk.Scrollbar(window, orient="horizontal", command=body.xview)
        body.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        body.pack(side="left", fill="both", expand=True)
        body.insert("1.0", text)
        body.configure(state="disabled")

    def _temporary(self) -> None:
        """Muestra dónde quedan los archivos temporales de la herramienta y permite borrarlos."""

        from tkinter import messagebox

        items = temporary_files()
        lines = [
            "ARCHIVOS TEMPORALES",
            "=" * 19,
            f"Conversiones de rwconverter (.rdf -> .xml): {TEMP_DIR}",
            "  Se borran solas al terminar «Nuevo proyecto desde DOCX»; el XML que el proyecto",
            "  necesita queda en generado/<nombre>.reports.xml (ese no es temporal).",
            f"Copias .bak y compilaciones interrumpidas: dentro de {default_projects_root()}",
            "",
        ]
        if not items:
            self._show("\n".join(lines + ["No hay archivos temporales."]))
            return
        total = sum(item.size for item in items)
        lines.append(f"Encontrados {len(items)} ({format_size(total)}):")
        lines += [f"  {item.path}\n      {item.kind}, {format_size(item.size)}" for item in items]
        self._show("\n".join(lines))
        if not messagebox.askyesno(
            "Borrar archivos temporales",
            f"¿Borrar los {len(items)} elementos listados ({format_size(total)})?\n\n"
            "Ningún proyecto deja de funcionar. Las copias .bak son la versión anterior de "
            "archivos reemplazados al recargar un DOCX: después no se podrán recuperar.",
            parent=self.root,
        ):
            return
        errors = remove_temporary(items)
        lines.append("")
        if errors:
            lines += ["No se pudieron borrar:"] + [f"  {error}" for error in errors]
        else:
            lines.append("Borrados.")
        self._show("\n".join(lines))

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
