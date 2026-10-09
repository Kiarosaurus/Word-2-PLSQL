"""Lectura del Modelo de Datos de un reporte de Oracle Reports.

``rwconverter`` (incluido en Reports Builder) convierte un ``.rdf`` a XML con
todo el Modelo de Datos: consultas con su SQL, grupos, columnas, fórmulas,
resúmenes, parámetros, Data Links y los campos del diseño con su Origen. Con
eso, en el Word basta escribir el nombre que se ve en Reports
(``{{FIELD:F_96}}``, ``{{COLUMN:INTERES2}}``, ``{{SUM:F_108}}``) y el
compilador averigua a qué consulta pertenece cada dato.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import difflib
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
from xml.etree import ElementTree as ET

from .diagnostics import Diagnostics


# Reports admite letras con tilde o «ñ» en los nombres (AÑO2); Oracle también.
IDENT = re.compile(r"^[^\W\d_][\w$#]{0,29}$")
RWCONVERTER_CANDIDATES = (
    r"D:\oracle\Middleware\FR_Home\bin\rwconverter.exe",
    r"C:\oracle\Middleware\FR_Home\bin\rwconverter.exe",
    r"D:\oracle\Middleware\Oracle_Home\bin\rwconverter.exe",
    r"C:\oracle\Middleware\Oracle_Home\bin\rwconverter.exe",
)


@dataclass(frozen=True, slots=True)
class ReportsColumn:
    name: str            # nombre en el Modelo de Datos, en mayúsculas
    query: str           # consulta a la que pertenece
    group: str | None
    kind: str            # column, formula, summary, placeholder
    source: str | None = None       # resumen: columna que acumula
    function: str | None = None     # resumen: sum, count...
    formula_code: str | None = None # fórmula: PL/SQL original
    datatype: str = "T"             # N número, D fecha, T texto
    unit: str | None = None         # fórmula: nombre de su función (program unit), en minúsculas
    reset: str | None = None        # resumen: grupo en el que vuelve a cero (o "report")


@dataclass(frozen=True, slots=True)
class Resolution:
    marker: str                  # CONSULTA.COLUMNA, o NOMBRE si es un parámetro
    mask: str | None = None      # máscara de formato del campo en Reports
    pending: str | None = None   # trabajo manual pendiente (fórmulas, resúmenes sueltos)


@dataclass(slots=True)
class ReportsModel:
    name: str
    queries: dict[str, dict] = field(default_factory=dict)      # CONSULTA -> {name, sql, columns}
    columns: dict[str, ReportsColumn] = field(default_factory=dict)
    fields: dict[str, tuple[str, str | None]] = field(default_factory=dict)  # F_x -> (origen, máscara)
    parameters: dict[str, dict] = field(default_factory=dict)
    links: list[dict] = field(default_factory=list)
    program_units: dict[str, str] = field(default_factory=dict)  # nombre (minúsculas) -> PL/SQL
    groups: dict[str, list[str]] = field(default_factory=dict)    # grupo (mayúsculas) -> columnas propias

    # ------------------------------------------------------------ resolución
    def suggestions(self, name: str) -> list[str]:
        pool = list(self.fields) + list(self.columns) + list(self.parameters)
        return difflib.get_close_matches(name.upper(), pool, n=3, cutoff=0.6)

    def resolve(self, kind: str, name: str) -> Resolution | None:
        """kind: FIELD, COLUMN o SUM. Devuelve None si el nombre no existe en Reports."""

        key = name.strip().upper()
        mask = None
        if key in self.fields:
            key, mask = self.fields[key]
        if key in self.parameters and kind == "FIELD":
            return Resolution(key, mask)
        column = self.columns.get(key)
        if column is None:
            return None
        if kind == "SUM":
            if column.kind == "summary":
                summed = self.columns.get(column.source or "")
                if column.function in (None, "sum") and summed is not None:
                    # El total de la tabla suma la columna origen mientras recorre las filas.
                    return Resolution(f"{summed.query}.{column.source}", mask)
                # Otras funciones (count, maximum...): el motor calcula el total CS_ directamente.
                return Resolution(f"{summed.query if summed else column.query}.{column.name}", mask)
            return Resolution(f"{column.query}.{column.name}", mask)
        # Fórmulas, marcadores de posición y totales los calcula el motor con el código
        # convertido (reports_convert); lo que no se pueda convertir se informa al compilar.
        return Resolution(f"{column.query}.{column.name}", mask)

    def query_of(self, name: str) -> str | None:
        resolution = self.resolve("COLUMN", name)
        return resolution.marker.split(".", 1)[0] if resolution and "." in resolution.marker else None


# ------------------------------------------------------------------ conversión
def find_rwconverter() -> Path | None:
    configured = os.environ.get("RWCONVERTER")
    candidates = ([configured] if configured else []) + list(RWCONVERTER_CANDIDATES)
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return Path(candidate)
    found = shutil.which("rwconverter")
    return Path(found) if found else None


def convert_rdf(rdf: Path, destination: Path, diagnostics: Diagnostics) -> Path | None:
    """Convierte un .rdf a XML con rwconverter (Reports Builder)."""

    tool = find_rwconverter()
    if tool is None:
        diagnostics.error(
            "REPORTS-001",
            "No se encontró rwconverter.exe de Oracle Reports para leer el .rdf.",
            suggestion="Defina la variable RWCONVERTER con su ruta, o convierta el .rdf a .xml en Reports Builder "
                       "(Archivo > Convertir) y use el .xml.",
        )
        return None
    try:
        subprocess.run(
            [str(tool), f"source={rdf}", f"dest={destination}", "stype=rdffile", "dtype=xmlfile",
             "batch=yes", "overwrite=yes"],
            check=False, capture_output=True, timeout=300,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        diagnostics.error("REPORTS-002", f"rwconverter no pudo convertir el .rdf: {exc}", location=str(rdf))
        return None
    if not destination.is_file() or destination.stat().st_size == 0:
        diagnostics.error("REPORTS-002", "rwconverter no generó el XML del reporte.", location=str(rdf))
        return None
    return destination


_CONVERTED: dict[tuple[str, int, int], Path] = {}
# Conversiones de rwconverter: dentro de la instalación (proyectos/_temporal), no en %TEMP%,
# para que el usuario vea dónde quedan y pueda borrarlas desde la herramienta.
TEMP_DIR = Path(__file__).resolve().parents[2] / "proyectos" / "_temporal"
RWCONVERTER_DIR = TEMP_DIR / "rwconverter"
LEGACY_TEMP_PREFIX = "reports-xml-"     # versiones anteriores convertían en %TEMP%\reports-xml-*


def discard_conversion(path: Path | None) -> None:
    """Borra el XML temporal convertido de ``path`` (.rdf) una vez copiado al proyecto."""

    if path is None:
        return
    resolved = str(Path(path).resolve()).lower()
    for key in [key for key in _CONVERTED if key[0] == resolved]:
        folder = _CONVERTED.pop(key).parent
        if folder.parent == RWCONVERTER_DIR:
            shutil.rmtree(folder, ignore_errors=True)


def reports_xml(path: Path, diagnostics: Diagnostics, work_dir: Path | None = None) -> Path | None:
    """Ruta a un XML del Modelo de Datos: el mismo archivo si es .xml, o el convertido si es .rdf.

    El .rdf se convierte en ``proyectos/_temporal/rwconverter/<nombre>-xxxx/`` (o en
    ``work_dir``); quien lo copie al proyecto debe llamar a :func:`discard_conversion`.
    """

    path = Path(path)
    if path.suffix.lower() == ".xml":
        return path if path.is_file() else None
    if path.suffix.lower() != ".rdf":
        diagnostics.error("REPORTS-003", "El reporte de Oracle debe ser un .rdf o un .xml.", location=str(path))
        return None
    if not path.is_file():
        diagnostics.error("REPORTS-003", "No existe el archivo del reporte de Oracle.", location=str(path))
        return None
    # rwconverter tarda varios segundos: se reutiliza la conversión del mismo .rdf sin cambios.
    stat = path.stat()
    key = (str(path.resolve()).lower(), stat.st_mtime_ns, stat.st_size)
    cached = _CONVERTED.get(key)
    if cached is not None and cached.is_file():
        return cached
    try:
        RWCONVERTER_DIR.mkdir(parents=True, exist_ok=True)
        target_dir = Path(work_dir) if work_dir else Path(tempfile.mkdtemp(prefix=f"{path.stem}-", dir=RWCONVERTER_DIR))
    except OSError as exc:
        diagnostics.error("REPORTS-002", f"No se pudo crear la carpeta temporal de la conversión: {exc}",
                          location=str(RWCONVERTER_DIR))
        return None
    converted = convert_rdf(path, target_dir / f"{path.stem}.xml", diagnostics)
    if converted is not None:
        _CONVERTED[key] = converted
    elif not work_dir:
        shutil.rmtree(target_dir, ignore_errors=True)
    return converted


# ------------------------------------------------------------------ lectura
def _ident(name: str | None) -> str | None:
    if not name:
        return None
    value = name.strip().upper()
    return value if IDENT.fullmatch(value) else None


def _datatype(value: str | None) -> str:
    value = (value or "").lower()
    if "number" in value or value in {"integer", "float"}:
        return "N"
    if "date" in value:
        return "D"
    return "T"


def _report_mask(mask: str | None) -> str | None:
    """Máscara de Reports -> máscara TO_CHAR (N = dígito con supresión de ceros)."""

    if not mask:
        return None
    if re.fullmatch(r"[N90,.\-$]+", mask):
        return mask.replace("N", "9")
    if re.fullmatch(r"[A-Za-z0-9/:. \-]+", mask):
        return mask
    return None


def load_reports_model(xml_path: Path, diagnostics: Diagnostics) -> ReportsModel | None:
    try:
        root = ET.parse(xml_path).getroot()
    except (OSError, ET.ParseError) as exc:
        diagnostics.error("REPORTS-004", f"No se pudo leer el XML de Reports: {exc}", location=str(xml_path))
        return None
    if root.tag != "report" or root.find("data") is None:
        diagnostics.error("REPORTS-004", "El XML no es un reporte de Oracle Reports.", location=str(xml_path))
        return None

    model = ReportsModel(root.get("name") or Path(xml_path).stem)
    functions = {
        (fn.get("name") or "").lower(): (fn.findtext("textSource") or "").strip()
        for fn in root.iter("function")
    }
    model.program_units = dict(functions)
    data = root.find("data")
    for source in data.findall("dataSource"):
        query = _ident(source.get("name"))
        if query is None:
            diagnostics.warning("REPORTS-005", f"Consulta con nombre no utilizable: {source.get('name')!r}.")
            continue
        items = []
        for group in source.iter("group"):
            group_name = group.get("name")
            model.groups[(group_name or "").upper()] = [
                _ident(child.get("name")) for child in group if child.tag == "dataItem" and _ident(child.get("name"))
            ]
            for child in group:
                name = _ident(child.get("name"))
                if child.tag == "dataItem":
                    # La posición se conserva aunque el nombre no sea utilizable (por ejemplo
                    # con «ñ»): el motor nombra las columnas por posición en el SELECT.
                    descriptor = child.find("dataDescriptor")
                    order = int(descriptor.get("order")) if descriptor is not None and descriptor.get("order") else 0
                    items.append((order, name))
                if name is None:
                    continue
                if child.tag == "dataItem":
                    descriptor = child.find("dataDescriptor")
                    kind = child.get("oracleDatatype") or (descriptor.get("oracleDatatype") if descriptor is not None else None)
                    model.columns[name] = ReportsColumn(name, query, group_name, "column", datatype=_datatype(kind))
                elif child.tag == "formula":
                    code = functions.get((child.get("source") or "").lower())
                    model.columns[name] = ReportsColumn(name, query, group_name, "formula", formula_code=code,
                                                        datatype=_datatype(child.get("datatype")),
                                                        unit=(child.get("source") or "").lower() or None)
                elif child.tag == "placeholder":
                    model.columns[name] = ReportsColumn(name, query, group_name, "placeholder",
                                                        datatype=_datatype(child.get("datatype")))
                elif child.tag == "summary":
                    model.columns[name] = ReportsColumn(
                        name, query, group_name, "summary", datatype="N",
                        source=_ident(child.get("source")), function=(child.get("function") or "").lower() or None,
                        reset=(child.get("reset") or "report").upper())
        model.queries[query] = {
            "name": source.get("name"),
            "sql": (source.findtext("select") or "").strip(),
            "columns": [name for _order, name in sorted(items)],
        }
    # Resúmenes y fórmulas a nivel de reporte: pertenecen a la consulta de su columna origen.
    for child in data:
        name = _ident(child.get("name"))
        if name is None or child.tag not in {"summary", "formula", "placeholder"}:
            continue
        summed = _ident(child.get("source")) if child.tag == "summary" else None
        owner = model.columns.get(summed or "")
        query = owner.query if owner else next(iter(model.queries), "REPORTE")
        model.columns[name] = ReportsColumn(
            name, query, None, child.tag, source=summed,
            function=(child.get("function") or "").lower() or None,
            formula_code=functions.get((child.get("source") or "").lower()) if child.tag == "formula" else None,
            datatype="N" if child.tag == "summary" else _datatype(child.get("datatype")),
            unit=(child.get("source") or "").lower() or None if child.tag == "formula" else None,
            reset=(child.get("reset") or "report").upper() if child.tag == "summary" else None)
    for parameter in data.findall("userParameter"):
        name = _ident(parameter.get("name"))
        if name:
            model.parameters[name] = {"datatype": parameter.get("datatype") or "character",
                                      "initial": parameter.get("initialValue")}
    for parameter in data.findall("systemParameter"):
        name = _ident(parameter.get("name"))
        if name and name not in model.parameters:
            model.parameters[name] = {"datatype": "character", "initial": parameter.get("initialValue"),
                                      "system": True}
    for link in data.findall("link"):
        model.links.append({key: link.get(key) for key in ("parentGroup", "parentColumn", "childQuery", "childColumn",
                                                            "condition", "sqlClause")})
    for item in root.iter("field"):
        name, source = _ident(item.get("name")), _ident(item.get("source"))
        if name and source:
            model.fields[name] = (source, _report_mask(item.get("formatMask")))
    return model
