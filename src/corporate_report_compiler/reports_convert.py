"""Conversión del código de Oracle Reports al package del reporte (modo layout).

A partir del Modelo de Datos (``reports_model``) y de los nombres usados en el
Word, arma el plan que el motor RPT_LAYOUT necesita para calcular lo mismo que
Reports:

- Fórmulas (CF_): su función PL/SQL se copia al package del reporte. Cada
  referencia ``:NOMBRE`` pasa a una variable local leída con ``rpt_layout.num``,
  ``txt`` o ``dat``; las asignaciones ``:CP_X := ...`` y ``INTO :CP_X`` se
  devuelven al motor con ``rpt_layout.set_num``, ``set_txt`` o ``set_dat``.
- Funciones auxiliares (program units) que usan las fórmulas: se copian igual.
- Marcadores de posición (CP_): toman el valor que les asignan las fórmulas.
- Totales (CS_): el motor los calcula recorriendo su consulta.
- Data Links: el motor filtra cada fila de la consulta hija con el valor de la
  columna del grupo padre.

Lo que no se puede convertir (SRW, nombres inexistentes) queda como función que
devuelve NULL con el código original comentado, y se informa como pendiente.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import re
import unicodedata

from .reports_model import IDENT, ReportsModel


TOKEN = re.compile(
    r"(?P<ws>\s+)"
    r"|(?P<comment>--[^\n]*|/\*.*?\*/)"
    r"|(?P<string>'(?:[^']|'')*')"
    r"|(?P<assign>:=)"
    r"|(?P<bind>:[^\s:=;,()+\-*/|<>!'\"%.]+)"
    r"|(?P<ident>[^\W\d][\w$#]*)"
    r"|(?P<other>.)",
    re.DOTALL,
)
CODE_DIR = "codigo_reports"     # funciones de Reports corregidas a mano (generado/codigo_reports)
GETTER = {"N": "num", "D": "dat", "T": "txt"}
LOCAL_TYPE = {"N": "NUMBER", "D": "DATE", "T": "VARCHAR2(4000)"}
CONDITIONS = {"eq": "eq", "ne": "ne", "notequal": "ne", "lt": "lt", "le": "le", "gt": "gt", "ge": "ge",
              "lessthan": "lt", "lessthanorequal": "le", "greaterthan": "gt", "greaterthanorequal": "ge"}
SUMMARY_FUNCTIONS = {"sum", "count", "minimum", "maximum", "average", "first", "last"}
# Funciones de SQL/PL/SQL y palabras que van seguidas de «(»: no son funciones de la base de datos.
BUILTINS = set("""
abs add_months and ascii avg between case cast ceil char chr coalesce concat count current_date decode
distinct else elsif exists extract floor from greatest if in initcap instr is last_day least length like
listagg lower lpad ltrim max min mod months_between not number nvl nvl2 or power raise_application_error
rank replace return round row_number rpad rtrim select sign sqrt substr sum sysdate systimestamp then
to_char to_date to_number to_timestamp translate trim trunc upper using values varchar varchar2 when where
while date integer raw float numeric decimal regexp_replace regexp_substr regexp_like regexp_instr any all
nchar nvarchar2 sqlerrm sqlcode open fetch close loop for into on set update delete insert merge exit
""".split())


@dataclass(slots=True)
class Unit:
    name: str          # nombre de la función en minúsculas
    spec: str          # declaración para la especificación del package
    body: str          # código convertido
    stub: str | None = None   # motivo si no se pudo convertir


@dataclass(slots=True)
class ReportsPlan:
    formulas: dict[str, dict] = field(default_factory=dict)       # CF -> {q, t, f}
    placeholders: dict[str, dict] = field(default_factory=dict)   # CP -> {t, by}
    summaries: dict[str, dict] = field(default_factory=dict)      # CS -> {q, s, f[, run, each, rk]}
    filters: dict[str, list[dict]] = field(default_factory=dict)  # consulta hija -> [{c, op, p}]
    query_formulas: dict[str, list[str]] = field(default_factory=dict)
    query_refs: dict[str, list[str]] = field(default_factory=dict)  # consulta -> :COLUMNAS de otro grupo
    units: list[Unit] = field(default_factory=list)
    queries: set[str] = field(default_factory=set)                # consultas necesarias
    parameters: set[str] = field(default_factory=set)             # parámetros que usa el código
    pending: list[str] = field(default_factory=list)              # trabajo manual
    externals: set[str] = field(default_factory=set)              # funciones de la base de datos

    @property
    def empty(self) -> bool:
        return not (self.formulas or self.placeholders or self.summaries or self.filters)

    def engine_model(self, package: str) -> dict:
        """Lo que RPT_LAYOUT recibe en p_model."""

        return {"package": package.upper(), "formulas": self.formulas,
                "placeholders": self.placeholders, "summaries": self.summaries}


def tokens(code: str) -> list[tuple[str, str]]:
    return [(match.lastgroup, match.group()) for match in TOKEN.finditer(code)]


def _significant(items: list[tuple[str, str]], start: int, step: int = 1) -> int | None:
    index = start
    while 0 <= index < len(items):
        if items[index][0] not in {"ws", "comment"}:
            return index
        index += step
    return None


def local_name(ref: str) -> str:
    """Variable local para :REF, sin tildes ni «ñ»."""

    plain = unicodedata.normalize("NFKD", ref).encode("ascii", "ignore").decode()
    return "rr_" + plain.lower()


def _commented(code: str) -> str:
    return "\n".join(f"    --   {line.rstrip()}" for line in code.strip().splitlines())


def _datatype_of(model: ReportsModel, name: str) -> str | None:
    if name in model.columns:
        return model.columns[name].datatype
    if name in model.parameters:
        kind = (model.parameters[name].get("datatype") or "").lower()
        return "N" if "number" in kind else "D" if "date" in kind else "T"
    return None


@dataclass(slots=True)
class _Converted:
    unit: Unit
    reads: set[str]
    assigns: set[str]
    calls: set[str]
    externals: set[str]


def convert_unit(name: str, code: str, model: ReportsModel, *, datatype: str = "T") -> _Converted:
    """Convierte una función de Reports a PL/SQL que usa el motor en vez de las referencias :X."""

    items = tokens(code)
    reads: set[str] = set()
    assigns: set[str] = set()
    calls: set[str] = set()
    externals: set[str] = set()
    problems: list[str] = []

    # Cabecera: desde FUNCTION/PROCEDURE hasta el IS/AS de nivel 0.
    start = next((i for i, (kind, text) in enumerate(items)
                  if kind == "ident" and text.lower() in {"function", "procedure"}), None)
    header_end = None
    if start is not None:
        depth = 0
        for i in range(start, len(items)):
            kind, text = items[i]
            if text == "(":
                depth += 1
            elif text == ")":
                depth -= 1
            elif kind == "ident" and depth == 0 and text.lower() in {"is", "as"}:
                header_end = i
                break
    if header_end is None:
        problems.append("no tiene la forma «function ... return ... is»")

    # Referencias y llamadas.
    previous = None
    for i, (kind, text) in enumerate(items):
        if kind == "bind":
            ref = text[1:].upper()
            if not IDENT.fullmatch(ref):
                problems.append(f"usa :{text[1:]}, un nombre que el motor no puede leer")
            elif _datatype_of(model, ref) is None:
                problems.append(f"usa :{ref}, que no existe en el Modelo de Datos")
        elif kind == "ident" and header_end is not None and i > header_end:
            lower = text.lower()
            following = _significant(items, i + 1)
            next_text = items[following][1] if following is not None else ""
            if lower == "srw" and next_text == ".":
                problems.append("usa el paquete SRW de Oracle Reports")
            elif previous != "." and lower in model.program_units and lower != name:
                calls.add(lower)
            elif previous != "." and next_text == "." and lower not in BUILTINS:
                after = _significant(items, following + 1)
                member = items[after][1] if after is not None and items[after][0] == "ident" else ""
                paren = _significant(items, after + 1) if after is not None else None
                if member and paren is not None and items[paren][1] == "(" and lower not in {"dbms_output"}:
                    externals.add(f"{text}.{member}".upper())
            elif previous != "." and next_text == "(" and lower not in BUILTINS:
                externals.add(text.upper())
        if kind not in {"ws", "comment"}:
            previous = text
    if problems:
        header = f"FUNCTION {name} RETURN {LOCAL_TYPE[datatype].split('(')[0]}"
        if header_end is not None:
            header = " ".join("".join(text for kind, text in items[start:header_end] if kind != "comment").split())
        reason = "; ".join(dict.fromkeys(problems))
        body = "\n".join([
            f"{header} IS",
            "BEGIN",
            f"    -- PENDIENTE: no se convirtió porque {reason}. Código original de Oracle Reports:",
            _commented(code),
            "    RETURN NULL;" if header.lower().startswith("function") else "    NULL;",
            f"END {name};",
        ])
        return _Converted(Unit(name, header + ";", body, reason), set(), set(), set(), set())

    # Conversión: :X -> variable local; asignaciones devueltas al motor al terminar la sentencia.
    out: list[str] = []
    to_set: list[str] = []
    in_into = False
    for i, (kind, text) in enumerate(items):
        if i <= header_end:
            out.append(text)
            continue
        lowered = text.lower() if kind == "ident" else ""
        if lowered == "into":
            in_into = True
        elif lowered in {"from", "using", "bulk", "values", "select"} or text == ";":
            in_into = False
        if kind == "bind":
            ref = text[1:].upper()
            following = _significant(items, i + 1)
            assigned = in_into or (following is not None and items[following][0] == "assign")
            if assigned:
                assigns.add(ref)
                if ref not in to_set:
                    to_set.append(ref)
            else:
                reads.add(ref)
            out.append(local_name(ref))
            continue
        out.append(text)
        if text == ";" and to_set:
            setters = " ".join(
                f"rpt_layout.set_{GETTER[_datatype_of(model, ref)]}('{ref}', {local_name(ref)});" for ref in to_set)
            out.append(" " + setters)
            to_set = []
    declarations = []
    for ref in sorted(reads | assigns):
        kind = _datatype_of(model, ref)
        init = f" := rpt_layout.{GETTER[kind]}('{ref}')" if ref in reads else ""
        declarations.append(f"\n    {local_name(ref)} {LOCAL_TYPE[kind]}{init};")
    out.insert(header_end + 1, "".join(declarations))
    header = " ".join("".join(text for kind, text in items[start:header_end] if kind != "comment").split())
    body = "".join(out[start:]).strip()
    if not body.endswith(";"):
        body += ";"
    return _Converted(Unit(name, header + ";", body), reads, assigns, calls, externals)


def plan_conversion(model: ReportsModel, used: set[str]) -> ReportsPlan:
    """Arma el plan para los nombres de Reports que usa el Word (columnas, fórmulas, totales...)."""

    plan = ReportsPlan()
    converted: dict[str, _Converted] = {}

    def convert(unit: str, datatype: str = "T") -> _Converted:
        if unit not in converted:
            converted[unit] = convert_unit(unit, model.program_units.get(unit, ""), model, datatype=datatype)
        return converted[unit]

    # Quién asigna cada marcador de posición (para calcularlo antes de leerlo).
    assigners: dict[str, list[str]] = {}
    for column in model.columns.values():
        if column.kind == "formula" and column.unit in model.program_units:
            for target in convert(column.unit, column.datatype).assigns:
                assigners.setdefault(target, []).append(column.name)

    pending = sorted(used, reverse=True)    # orden estable: el package generado no cambia entre corridas
    seen: set[str] = set()
    units: list[str] = []

    def add_unit(unit: str, datatype: str = "T") -> None:
        if unit in units:
            return
        units.append(unit)
        result = convert(unit, datatype)
        pending.extend(sorted(result.reads | result.assigns, reverse=True))
        plan.externals |= result.externals
        for call in sorted(result.calls):
            add_unit(call)

    while pending:
        name = pending.pop()
        if name in seen:
            continue
        seen.add(name)
        if name in model.parameters:
            plan.parameters.add(name)
            continue
        column = model.columns.get(name)
        if column is None:
            continue
        if column.kind == "column":
            plan.queries.add(column.query)
        elif column.kind == "formula":
            if not column.unit or column.unit not in model.program_units:
                plan.pending.append(f"La fórmula {name} no tiene código en el .rdf.")
                continue
            plan.formulas[name] = {"q": column.query if column.group else None, "t": column.datatype,
                                   "f": column.unit}
            if column.group:
                plan.queries.add(column.query)
            add_unit(column.unit, column.datatype)
        elif column.kind == "placeholder":
            plan.placeholders[name] = {"t": column.datatype, "by": assigners.get(name, [])}
            pending.extend(assigners.get(name, []))
            if not assigners.get(name):
                plan.pending.append(f"{name} es un marcador de posición que ninguna fórmula asigna: quedará vacío.")
        elif column.kind == "summary":
            source = model.columns.get(column.source or "")
            function = (column.function or "sum").lower()
            if source is None:
                plan.pending.append(f"El total {name} suma {column.source}, que no está en el Modelo de Datos.")
                continue
            if function not in SUMMARY_FUNCTIONS:
                plan.pending.append(f"El total {name} usa la función {function}, que el motor no calcula.")
                continue
            plan.summaries[name] = {"q": source.query, "s": source.name, "f": function}
            if column.group and source.group and column.group.upper() == source.group.upper():
                # Total en el mismo grupo que su origen: en Reports es acumulado fila a fila
                # (por ejemplo la numeración 1, 2, 3 con sum de una columna que vale 1).
                reset = (column.reset or "REPORT").upper()
                plan.summaries[name]["run"] = True
                if reset == column.group.upper():
                    plan.summaries[name]["each"] = True
                elif reset != "REPORT" and set(model.groups.get(reset, [])) and all(
                        model.columns.get(c) is not None and model.columns[c].query == source.query
                        for c in model.groups[reset]):
                    plan.summaries[name]["rk"] = model.groups[reset]   # vuelve a cero al cambiar este grupo
            plan.queries.add(source.query)
            pending.append(source.name)
        # Data Links y referencias :COLUMNA de las consultas que se van agregando.
        for query in sorted(plan.queries - set(plan.filters)):
            plan.filters[query] = []
            sql = model.queries.get(query, {}).get("sql", "")
            for kind, text in tokens(sql):
                ref = text[1:].upper() if kind == "bind" else None
                if ref and ref in model.parameters:
                    plan.parameters.add(ref)
                elif ref and ref in model.columns and model.columns[ref].query != query:
                    # En Reports, :COLUMNA en el SQL toma el valor de la fila actual de otro grupo.
                    plan.query_refs.setdefault(query, [])
                    if ref not in plan.query_refs[query]:
                        plan.query_refs[query].append(ref)
                    pending.append(ref)
            for link in model.links:
                if (link.get("childQuery") or "").upper() != query:
                    continue
                child = (link.get("childColumn") or "").upper()
                parent = (link.get("parentColumn") or "").upper()
                condition = CONDITIONS.get((link.get("condition") or "eq").lower())
                if child not in model.columns or parent not in model.columns or condition is None:
                    plan.pending.append(
                        f"Data Link {link.get('parentGroup')}.{link.get('parentColumn')} -> {query}.{link.get('childColumn')} "
                        f"({link.get('condition')}): agregue la condición a mano en q_{query.lower()}.sql.")
                    continue
                plan.filters[query].append({"c": child, "op": condition, "p": parent})
                pending.append(parent)
    plan.filters = {query: filters for query, filters in sorted(plan.filters.items()) if filters}
    plan.formulas = dict(sorted(plan.formulas.items()))
    plan.placeholders = dict(sorted(plan.placeholders.items()))
    plan.summaries = dict(sorted(plan.summaries.items()))

    for column in model.columns.values():            # orden del Modelo de Datos
        entry = plan.formulas.get(column.name)
        if entry and entry["q"]:
            plan.query_formulas.setdefault(entry["q"], []).append(column.name)
    for unit in sorted(units):
        result = converted[unit]
        plan.units.append(result.unit)
        if result.unit.stub:
            owner = next((c.name for c in model.columns.values() if c.unit == unit), unit)
            plan.pending.append(f"{owner}: no se convirtió porque {result.unit.stub}. Corrija su código en "
                                f"generado/{CODE_DIR}/{unit}.sql (con la sintaxis de Reports) y recompile; "
                                "mientras tanto devuelve NULL.")
    plan.externals -= {unit.upper() for unit in model.program_units}
    plan.pending = list(dict.fromkeys(plan.pending))
    return plan
