from __future__ import annotations

import os
from pathlib import Path
import re
import sys
import traceback

from .apex_guide import (
    PAGE_PLACEHOLDER,
    build_apex_guide,
    build_detailed_instructions,
    build_paste_instructions,
    process_code,
)
from .compiler import CompilationResult, compile_project, validate_project
from .workspace import (
    change_page,
    current_page,
    page_items,
    rename_items,
    default_projects_root,
    existing_page,
    files_to_replace,
    format_import,
    import_docx,
    format_size,
    format_update,
    page_placeholder_locations,
    remove_temporary,
    temporary_files,
    update_docx,
    workspace_outputs,
)
from .reports_model import TEMP_DIR, discard_conversion
from .sources import common_source, project_tables, qualified, set_table_source


def _project_location(item: dict, project_path: Path | None) -> str | None:
    """Archivo (y línea, si el mensaje cita un valor) de los errores del .report.json sin ubicación."""

    if project_path is None or not re.match(r"(LAYOUT-)?PROJECT-", item["code"]):
        return None
    path = Path(project_path)
    layout = workspace_outputs(path)
    shown = path.relative_to(layout[1]).as_posix() if layout is not None else str(path)
    quoted = re.search(r"'([^']+)'", str(item["message"]))
    if quoted:
        try:
            text = path.read_text(encoding="utf-8-sig")
        except OSError:
            return shown
        for number, line in enumerate(text.splitlines(), 1):
            if quoted.group(1) in line:
                return f"{shown}:{number}"
    return shown


def _format_diagnostics(items: list[dict], project_path: Path | None = None) -> list[str]:
    lines: list[str] = []
    for item in items:
        where = item.get("location") or _project_location(item, project_path)
        location = f" — {where}" if where else ""
        lines.append(f"{item['severity']} {item['code']}{location}")
        lines.append(str(item["message"]))
        if "suggestion" in item:
            lines.append(f"Solución: {item['suggestion']}")
        lines.append("")
    return lines


def bold_lines(text: str) -> list[int]:
    """Líneas (desde 1) que la GUI muestra en negrita: los errores y dónde arreglarlos.

    - Cada bloque «ERROR ...» hasta la línea en blanco: código, archivo, mensaje y «Solución:».
    - «Reemplace XX ...» con la lista archivo:línea que le sigue.
    - El resumen «No válido: ...».
    """

    bold: list[int] = []
    block = None                     # "error" o "locations"
    for number, line in enumerate(text.splitlines(), 1):
        if line.startswith("ERROR "):
            block = "error"
        elif line.startswith(f"Reemplace {PAGE_PLACEHOLDER} por el número de página en:"):
            block = "locations"
        elif block == "error" and not line.strip():
            block = None
        elif block == "locations" and not line.startswith("  "):
            block = None
        elif block == "locations" and line.startswith("  ("):
            continue                 # nota aclaratoria, no un lugar que corregir
        if block is not None or line.startswith("No válido:"):
            bold.append(number)
    return bold


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
                     "por la página real con «Administrar items y tablas…»).")
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
             + _format_diagnostics(result.diagnostics.to_dict()["diagnostics"], result.project_path))
    if result.valid:
        if compiled:
            lines.append("Compilación correcta. Archivos generados:")
            lines.extend(f"  {path}" for path in result.artifacts)
            lines.append("")
            lines.append(build_apex_guide(result))
        else:
            lines.append("Validación correcta; no se escribió ningún archivo.")
            lines.append("Pulse Compilar para generar o actualizar los archivos citados abajo. «Copiar código APEX»")
            lines.append("ya copia el proceso y abre dónde pegarlo, con acceso a las instrucciones detalladas.")
            lines.append("")
            lines.append(build_apex_guide(result))
    elif not lines:
        lines.append("La validación falló sin diagnóstico disponible.")
    else:
        lines.append(
            f"No válido: {result.diagnostics.error_count} error(es), "
            f"{result.diagnostics.warning_count} advertencia(s)."
        )
        lines.append("«Copiar código APEX» queda desactivado mientras haya errores: el proceso se genera "
                     "solo con un proyecto válido.")
        lines.append("Al corregirlos, Validar y Compilar muestran qué subir a APEX y dónde.")
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


def ask_page(parent, title: str, prompt: str, *, initial: str = "") -> str | None:
    """Como ``simpledialog.askstring``, con «XX» en gris de fondo mientras el campo esté vacío.

    Devuelve el texto escrito (vacío si no se escribió nada) o None si se cancela.
    """

    import tkinter as tk
    from tkinter import simpledialog, ttk

    class PageDialog(simpledialog.Dialog):
        value: str | None = None

        def body(self, master):
            ttk.Label(master, text=prompt, justify="left").grid(row=0, column=0, sticky="w", padx=6, pady=(6, 4))
            self.text = tk.StringVar(value=initial)
            entry = ttk.Entry(master, textvariable=self.text, width=12)
            entry.grid(row=1, column=0, sticky="w", padx=6, pady=(0, 6))
            hint = tk.Label(entry, text=PAGE_PLACEHOLDER, fg="gray60", bg="white", cursor="xterm")
            hint.bind("<Button-1>", lambda _event: entry.focus_set())

            def refresh(*_args) -> None:
                if self.text.get():
                    hint.place_forget()
                else:
                    hint.place(x=3, rely=0.5, anchor="w")

            self.text.trace_add("write", refresh)
            refresh()
            return entry

        def apply(self):
            self.value = self.text.get().strip().upper()

    return PageDialog(parent, title).value


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
        self._last_result: CompilationResult | None = None
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
        ttk.Button(actions, text="Actualizar Word…", command=self._update_docx).pack(side="left", padx=(0, 16))
        ttk.Button(actions, text="Validar", command=self._validate).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Compilar", command=self._compile).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Administrar items y tablas…", command=self._manage_project).pack(side="left", padx=(0, 16))
        self.copy_button = ttk.Button(actions, text="Copiar código APEX", command=self._copy_process, state="disabled")
        self.copy_button.pack(side="left", padx=(0, 8))
        self.open_button = ttk.Button(actions, text="Abrir carpeta de salida", command=self._open_output, state="disabled")
        self.open_button.pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Archivos temporales…", command=self._temporary).pack(side="left")

        text_frame = ttk.Frame(frame)
        text_frame.grid(row=4, column=0, columnspan=3, sticky="nsew")
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)
        self.result = tk.Text(text_frame, wrap="none", state="disabled", font=("Consolas", 9))
        self.result.tag_configure("bold", font=("Consolas", 9, "bold"))
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
        for number in bold_lines(value):
            self.result.tag_add("bold", f"{number}.0", f"{number}.end")
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
        from tkinter import filedialog

        value = filedialog.askopenfilename(
            title="Seleccione la plantilla Word",
            filetypes=(("Documento de Word", "*.docx"), ("Todos", "*.*")),
        )
        if not value:
            return
        previous = existing_page(Path(value))
        page = ask_page(
            self.root,
            "Número de página APEX",
            "Número de la página APEX que tendrá los filtros.\n"
            f"Déjelo vacío para generar items P{PAGE_PLACEHOLDER}_ y cambiarlo después con «Administrar items y tablas…»."
            + (f"\n\nEste DOCX ya es un proyecto con la página {previous}: se propone la misma "
               f"para no volver a P{PAGE_PLACEHOLDER}_." if previous else ""),
            initial=previous or "",
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

    def _update_docx(self) -> None:
        """Reemplaza solo el Word del proyecto; conserva Page Items, tablas y demás ajustes."""

        from tkinter import filedialog, messagebox

        path = self._path_or_message()
        if path is None:
            return
        value = filedialog.askopenfilename(
            title="Seleccione el Word actualizado",
            initialdir=str(path.parent.parent),
            filetypes=(("Documento de Word", "*.docx"), ("Todos", "*.*")),
        )
        if not value:
            return
        if not messagebox.askyesno(
            "Actualizar Word",
            "Se reemplaza solo la plantilla Word del proyecto (la anterior queda como .bak).\n\n"
            "Se conservan los Page Items, el título, los q_*.sql con sus tablas y las fórmulas corregidas. "
            "Si el Word nuevo usa datos que el proyecto no tenía, se agregan.\n\n¿Continuar?",
            parent=self.root,
        ):
            return

        def action() -> None:
            update = update_docx(path, Path(value))
            if update.docx is None:
                lines = ["NO SE ACTUALIZÓ EL WORD (el proyecto no cambió)", "=" * 46, ""]
                self._show("\n".join(lines + _format_diagnostics(update.diagnostics.to_dict()["diagnostics"])))
                return
            result = validate_project(path, strict=self.strict.get())
            self._show("\n".join(format_update(update)) + _format_result(result))
            self._remember_process(result)

        self._run(action)

    def _remember_process(self, result: CompilationResult) -> None:
        """Deja listo «Copiar código APEX» (y la carpeta) tanto al validar como al compilar."""

        from .apex_guide import expected_artifact

        self._process_text = process_code(result)
        self._last_result = result
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
        if self._last_result is not None:
            self._text_window("Dónde pegar el código APEX",
                              build_paste_instructions(self._last_result, self._process_text), size=(860, 520),
                              extra=("Instrucciones APEX detalladas", self._instructions))

    def _instructions(self) -> None:
        """Ventana con los pasos detallados de APEX para el proyecto actual (simple o layout)."""

        path = self._path_or_message()
        if path is None:
            return
        result = validate_project(path, strict=self.strict.get())
        self._remember_process(result)
        self._text_window("Instrucciones APEX detalladas", build_detailed_instructions(result))

    def _text_window(self, title: str, text: str, *, size: tuple[int, int] = (860, 560),
                     extra: tuple[str, object] | None = None) -> None:
        """Ventana de solo lectura con «Copiar código APEX» y «Copiar instrucciones»."""

        window = self.tk.Toplevel(self.root)
        window.title(title)
        window.minsize(*size)
        bar = self.ttk.Frame(window, padding=(10, 8))
        bar.pack(side="top", fill="x")

        def copy(value: str) -> None:
            window.clipboard_clear()
            window.clipboard_append(value)

        self.ttk.Button(bar, text="Copiar código APEX", command=lambda: copy(self._process_text or ""),
                        state="normal" if self._process_text else "disabled").pack(side="left", padx=(0, 8))
        self.ttk.Button(bar, text="Copiar instrucciones", command=lambda: copy(text)).pack(side="left")
        if extra is not None:
            self.ttk.Button(bar, text=extra[0], command=extra[1]).pack(side="left", padx=(8, 0))
        body = self.tk.Text(window, wrap="none", font=("Consolas", 10))
        body.tag_configure("bold", font=("Consolas", 10, "bold"))
        scroll_y = self.ttk.Scrollbar(window, orient="vertical", command=body.yview)
        scroll_x = self.ttk.Scrollbar(window, orient="horizontal", command=body.xview)
        body.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        body.pack(side="left", fill="both", expand=True)
        body.insert("1.0", text)
        lines = text.splitlines()
        for number, line in enumerate(lines, 1):
            # Títulos (la línea siguiente es un subrayado de «=» o «-») en negrita.
            if number < len(lines) and lines[number] and set(lines[number]) <= {"=", "-"}:
                body.tag_add("bold", f"{number}.0", f"{number}.end")
        body.configure(state="disabled")

    def _manage_project(self) -> None:
        """Ventana «Administrar items y tablas»: pestañas Page Items y Tablas (FROM)."""

        path = self._path_or_message()
        if path is None:
            return
        tk, ttk = self.tk, self.ttk
        window = tk.Toplevel(self.root)
        window.title(f"Administrar items y tablas — {path.name}")
        window.transient(self.root)
        window.minsize(900, 520)
        notebook = ttk.Notebook(window)
        notebook.pack(fill="both", expand=True, padx=8, pady=8)
        items_frame = ttk.Frame(notebook, padding=12)
        tables_frame = ttk.Frame(notebook, padding=12)
        notebook.add(items_frame, text="  Page Items  ")
        notebook.add(tables_frame, text="  Tablas (FROM)  ")
        self._items_tab(items_frame, path, window)
        self._tables_tab(tables_frame, path, window)

    def _items_tab(self, frame, path: Path, window) -> None:
        """Pestaña Page Items: todos a una página (P<n>_) o uno por uno (avanzado)."""

        tk, ttk = self.tk, self.ttk
        items = page_items(path)
        if not items:
            ttk.Label(frame, text="El proyecto no declara Page Items (parámetros, campos ITEM, formato u "
                                  "orientación).").pack(anchor="w")
            return
        old = current_page(path)
        mode = tk.StringVar(value="page")
        modes = ttk.Frame(frame)
        modes.pack(anchor="w", pady=(8, 8))
        ttk.Radiobutton(modes, text="Todos en una página (P<n>_)", variable=mode, value="page").pack(side="left")
        ttk.Radiobutton(modes, text="Avanzado: uno por uno", variable=mode, value="each").pack(side="left", padx=(16, 0))

        # Todos en una página: un solo número reemplaza la página de cada item (salvo los de la página 0).
        simple = ttk.Frame(frame)
        ttk.Label(simple, text=f"Página actual: {old or 'ninguna'}"
                               + (" (sin número todavía)" if old == PAGE_PLACEHOLDER else "")).grid(
            row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(simple, text="Nueva página:  P").grid(row=1, column=0, sticky="w", pady=6)
        page = tk.StringVar(value="" if old in (None, PAGE_PLACEHOLDER) else old)
        page_entry = ttk.Entry(simple, textvariable=page, width=8)
        page_entry.grid(row=1, column=1, sticky="w")
        ttk.Label(simple, text="_<NOMBRE>").grid(row=1, column=2, sticky="w")
        preview = tk.Text(simple, height=min(len(items), 12) + 1, width=90, font=("Consolas", 9), state="disabled")
        preview.grid(row=2, column=0, columnspan=3, sticky="ew")
        prefix = f"P{old}_" if old else None

        def refresh(*_args) -> None:
            number = page.get().strip().upper() or "?"
            width = max(len(item.item) for item in items)
            lines = []
            for item in items:
                if prefix and item.item.startswith(prefix):
                    target = f"P{number}_{item.item[len(prefix):]}"
                    lines.append(f"{item.item:<{width}}  ->  {target:<{width + 4}}  {item.use}")
                else:
                    lines.append(f"{item.item:<{width}}      (no cambia)  {item.use}")
            preview.configure(state="normal")
            preview.delete("1.0", "end")
            preview.insert("1.0", "\n".join(lines))
            preview.configure(state="disabled")

        page.trace_add("write", refresh)
        refresh()

        # Avanzado: un campo por item con su nombre nuevo.
        advanced = ttk.Frame(frame)
        for column, title in enumerate(("Page Item actual", "Nuevo nombre", "Uso")):
            ttk.Label(advanced, text=title, font=("Segoe UI", 9, "bold")).grid(row=0, column=column, sticky="w",
                                                                              padx=(0, 12))
        names: dict[str, tk.StringVar] = {}
        for row, item in enumerate(dict.fromkeys(items), 1):
            if item.item in names:
                continue
            names[item.item] = tk.StringVar(value=item.item)
            ttk.Label(advanced, text=item.item, font=("Consolas", 9)).grid(row=row, column=0, sticky="w", padx=(0, 12))
            ttk.Entry(advanced, textvariable=names[item.item], width=34, font=("Consolas", 9)).grid(
                row=row, column=1, sticky="w", padx=(0, 12), pady=1)
            ttk.Label(advanced, text=item.use).grid(row=row, column=2, sticky="w")

        def switch(*_args) -> None:
            simple.pack_forget()
            advanced.pack_forget()
            (simple if mode.get() == "page" else advanced).pack(fill="both", expand=True)

        mode.trace_add("write", switch)
        switch()
        ttk.Label(frame, foreground="gray40", text=(
            "Se renombran en el .report.json y en los .sql del proyecto. Después cree o renombre los items "
            "en APEX y pulse Compilar.")).pack(anchor="w", pady=(8, 0))
        buttons = ttk.Frame(frame)
        buttons.pack(anchor="e", pady=(10, 0))

        def apply() -> None:
            if mode.get() == "page":
                change = change_page(path, page.get())
            else:
                change = rename_items(path, {old_name: value.get() for old_name, value in names.items()})
            if change.error:
                from tkinter import messagebox

                messagebox.showerror("Page Items", change.error, parent=window)
                return
            window.destroy()
            self._run(lambda: self._show_renamed(path, change))

        ttk.Button(buttons, text="Aplicar", command=apply).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="Cerrar", command=window.destroy).pack(side="left")
        page_entry.focus_set()

    def _tables_tab(self, frame, path: Path, window) -> None:
        """Pestaña Tablas (FROM): esquema y DB link de cada tabla de las consultas y fórmulas."""

        from tkinter import messagebox

        tk, ttk = self.tk, self.ttk
        found = project_tables(path)
        schema_guess, link_guess = common_source(path, found)
        ttk.Label(frame, text="Origen de las tablas en la base de datos", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(frame, foreground="gray30", justify="left", text=(
            "El SQL copiado de Oracle Reports nombra las tablas como las veía Reports (por ejemplo PGPAGO). Si en el\n"
            "parsing schema de APEX están en otro esquema o en otra base, escriba aquí el esquema y el DB link y\n"
            "aplíquelos a las tablas que lo necesiten. Deje un campo vacío si no hace falta.")).pack(anchor="w", pady=(2, 8))

        origin = ttk.Frame(frame)
        origin.pack(anchor="w", fill="x")
        schema = tk.StringVar(value=schema_guess)
        link = tk.StringVar(value=link_guess)
        ttk.Label(origin, text="Esquema:").grid(row=0, column=0, sticky="w")
        schema_entry = ttk.Entry(origin, textvariable=schema, width=18, font=("Consolas", 10))
        schema_entry.grid(row=0, column=1, sticky="w", padx=(4, 16))
        ttk.Label(origin, text="DB link:").grid(row=0, column=2, sticky="w")
        ttk.Entry(origin, textvariable=link, width=18, font=("Consolas", 10)).grid(row=0, column=3, sticky="w", padx=(4, 16))
        example = ttk.Label(origin, font=("Consolas", 10))
        example.grid(row=0, column=4, sticky="w")

        def target(table: str = "TABLA") -> str:
            return qualified(table, schema.get(), link.get())

        def show_example(*_args) -> None:
            example.configure(text=f"Quedará:  {target()}")

        schema.trace_add("write", show_example)
        link.trace_add("write", show_example)
        show_example()

        bulk = ttk.Frame(frame)
        bulk.pack(anchor="w", pady=(8, 8))
        status = ttk.Label(frame, foreground="#1b5e20")

        # Lista desplazable: una fila por tabla.
        holder = ttk.Frame(frame)
        holder.pack(fill="both", expand=True)
        canvas = tk.Canvas(holder, highlightthickness=0, borderwidth=0)
        scroll = ttk.Scrollbar(holder, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        rows = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=rows, anchor="nw")
        rows.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(
            int(-e.delta / 120), "units")))
        canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))

        def apply(tables: list[str] | None, remove: bool = False) -> None:
            change = set_table_source(path, tables, "" if remove else schema.get(), "" if remove else link.get())
            if change.error:
                messagebox.showinfo("Tablas", change.error, parent=window)
                return
            done = "sin esquema ni DB link" if remove else target("<tabla>")
            status.configure(text=f"Listo: {len(change.tables)} tabla(s) quedaron como {done}. Pulse Compilar y "
                                  "vuelva a ejecutar el package en la base.")
            lines = ["ORIGEN DE LAS TABLAS ACTUALIZADO", "=" * 32,
                     *(f"  {table} -> {qualified(table, '' if remove else schema.get(), '' if remove else link.get())}"
                       for table in change.tables),
                     "", "Archivos actualizados:",
                     *(f"  {file}  ({count} reemplazo(s))" for file, count in change.files),
                     "", "Siguiente: pulse Compilar y vuelva a ejecutar el package del reporte (rpt_*.sql) en la base."]
            self._show("\n".join(lines))
            fill()

        ttk.Button(bulk, text="Aplicar a todas", command=lambda: apply(None)).pack(side="left", padx=(0, 8))
        ttk.Button(bulk, text="Quitar esquema y DB link a todas",
                   command=lambda: apply(None, remove=True)).pack(side="left")

        def fill(tables: list | None = None) -> None:
            for child in rows.winfo_children():
                child.destroy()
            tables = tables if tables is not None else project_tables(path)
            if not tables:
                ttk.Label(rows, text="No se encontraron tablas en las consultas del proyecto.").grid(row=0, column=0)
                return
            for column, title in enumerate(("Tabla", "Ahora", "Se usa en", "")):
                ttk.Label(rows, text=title, font=("Segoe UI", 9, "bold")).grid(row=0, column=column, sticky="w",
                                                                              padx=(0, 14), pady=(0, 4))
            for row, table in enumerate(tables, 1):
                forms = table.forms
                now = "  /  ".join(f"{form} ({count})" if len(forms) > 1 else form for form, count in forms.items())
                places = table.places
                where = ", ".join(places[:3]) + (f" y {len(places) - 3} más" if len(places) > 3 else "")
                ttk.Label(rows, text=table.table, font=("Consolas", 9, "bold")).grid(row=row, column=0, sticky="w",
                                                                                   padx=(0, 14))
                ttk.Label(rows, text=now, font=("Consolas", 9)).grid(row=row, column=1, sticky="w", padx=(0, 14))
                ttk.Label(rows, text=where, foreground="gray30").grid(row=row, column=2, sticky="w", padx=(0, 14))
                actions = ttk.Frame(rows)
                actions.grid(row=row, column=3, sticky="w", pady=1)
                ttk.Button(actions, text="Aplicar", width=8,
                           command=lambda name=table.table: apply([name])).pack(side="left", padx=(0, 4))
                plain = all(form.upper() == table.table for form in forms)
                ttk.Button(actions, text="Quitar", width=7, state="disabled" if plain else "normal",
                           command=lambda name=table.table: apply([name], remove=True)).pack(side="left")

        fill(found)
        status.pack(anchor="w", pady=(8, 0))
        ttk.Label(frame, foreground="gray40", text=(
            "Se cambia en generado/q_*.sql y, para las fórmulas, en generado/codigo_reports/<función>.sql "
            "(reemplaza a la del .rdf).")).pack(anchor="w")
        ttk.Button(frame, text="Cerrar", command=window.destroy).pack(anchor="e", pady=(8, 0))

    def _show_renamed(self, path: Path, change) -> None:
        lines = ["PAGE ITEMS RENOMBRADOS", "=" * 22]
        lines += [f"  {old_item} -> {new_item}" for old_item, new_item in change.renamed.items()]
        lines += ["", "Archivos actualizados:"]
        lines += [f"  {target}  ({count} reemplazo(s))" for target, count in change.files]
        lines += ["", "Cree o renombre esos Page Items en la página de APEX y pulse Compilar para "
                      "regenerar los archivos generados.", "", ""]
        result = validate_project(path, strict=self.strict.get())
        self._show("\n".join(lines) + _format_result(result))
        self._remember_process(result)

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
