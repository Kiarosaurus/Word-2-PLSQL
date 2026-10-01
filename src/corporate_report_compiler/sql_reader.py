from __future__ import annotations

import re

from .diagnostics import Diagnostics
from .placeholders import IDENTIFIER_PATTERN


MAX_SQL_LENGTH = 32_767


def _mask_literals_and_comments(sql: str, diagnostics: Diagnostics) -> str:
    result = list(sql)
    index = 0
    state = "NORMAL"
    block_start = -1
    quoted_start = -1
    q_close = ""
    while index < len(sql):
        char = sql[index]
        nxt = sql[index + 1] if index + 1 < len(sql) else ""
        if state == "NORMAL":
            prefix_length = 0
            if char in "qQ" and nxt == "'":
                prefix_length = 2
            elif (
                char in "nN"
                and nxt in "qQ"
                and index + 2 < len(sql)
                and sql[index + 2] == "'"
            ):
                prefix_length = 3
            previous = sql[index - 1] if index else ""
            if prefix_length and previous and (previous.isalnum() or previous in "_$#"):
                prefix_length = 0

            if prefix_length:
                delimiter_index = index + prefix_length
                if delimiter_index >= len(sql):
                    diagnostics.error("SQL-008", "La consulta contiene un literal alternativo q sin delimitador.")
                    result[index] = " "
                    index += 1
                    continue
                delimiter = sql[delimiter_index]
                q_close = {"[": "]", "(": ")", "{": "}", "<": ">"}.get(delimiter, delimiter)
                quoted_start = index
                for position in range(index, delimiter_index + 1):
                    result[position] = " "
                index = delimiter_index
                state = "Q_STRING"
            elif char == "'":
                state = "STRING"
                result[index] = " "
            elif char == '"':
                state = "QUOTED_IDENTIFIER"
                result[index] = " "
            elif char == "-" and nxt == "-":
                state = "LINE_COMMENT"
                result[index] = result[index + 1] = " "
                index += 1
            elif char == "/" and nxt == "*":
                state = "BLOCK_COMMENT"
                block_start = index
                result[index] = result[index + 1] = " "
                index += 1
        elif state == "STRING":
            result[index] = " " if char not in "\r\n" else char
            if char == "'":
                if nxt == "'":
                    result[index + 1] = " "
                    index += 1
                else:
                    state = "NORMAL"
        elif state == "QUOTED_IDENTIFIER":
            result[index] = " " if char not in "\r\n" else char
            if char == '"':
                if nxt == '"':
                    result[index + 1] = " "
                    index += 1
                else:
                    state = "NORMAL"
        elif state == "LINE_COMMENT":
            if char in "\r\n":
                state = "NORMAL"
            else:
                result[index] = " "
        elif state == "BLOCK_COMMENT":
            result[index] = " " if char not in "\r\n" else char
            if char == "*" and nxt == "/":
                result[index + 1] = " "
                index += 1
                state = "NORMAL"
        elif state == "Q_STRING":
            result[index] = " " if char not in "\r\n" else char
            if char == q_close and nxt == "'":
                result[index + 1] = " "
                index += 1
                state = "NORMAL"
        index += 1

    if state == "STRING":
        diagnostics.error("SQL-003", "La consulta contiene un literal de texto sin cerrar.")
    elif state == "QUOTED_IDENTIFIER":
        diagnostics.error("SQL-004", "La consulta contiene un identificador entre comillas sin cerrar.")
    elif state == "BLOCK_COMMENT":
        diagnostics.error("SQL-005", "La consulta contiene un comentario de bloque sin cerrar.", location=f"carácter {block_start + 1}")
    elif state == "Q_STRING":
        diagnostics.error(
            "SQL-008",
            "La consulta contiene un literal alternativo q sin cerrar.",
            location=f"carácter {quoted_start + 1}",
        )
    return "".join(result)


def validate_sql(sql: str, diagnostics: Diagnostics) -> tuple[str, ...]:
    if not sql.strip():
        diagnostics.error("SQL-001", "La consulta SQL está vacía.")
        return ()
    sql_bytes = len(sql.encode("utf-8"))
    if sql_bytes > MAX_SQL_LENGTH:
        diagnostics.error("SQL-002", f"La consulta supera el límite de {MAX_SQL_LENGTH} bytes UTF-8.")
    if "\x00" in sql:
        diagnostics.error("SQL-009", "La consulta contiene un carácter NUL no permitido.")
    masked = _mask_literals_and_comments(sql, diagnostics)
    first = re.match(r"\s*(SELECT|WITH)\b", masked, flags=re.IGNORECASE)
    if not first:
        diagnostics.error("SQL-006", "La consulta debe comenzar con SELECT o WITH.")
    if ";" in masked:
        diagnostics.error("SQL-007", "No se permiten puntos y coma ni varias sentencias SQL.")
    if re.search(r"(?m)^\s*/\s*$", masked):
        diagnostics.error("SQL-010", "No se permite el terminador SQL*Plus '/'.")
    binds: list[str] = []
    for match in re.finditer(r"(?<!:):([A-Za-z][A-Za-z0-9_$#]*)(?![A-Za-z0-9_$#])", masked):
        name = match.group(1).upper()
        if not IDENTIFIER_PATTERN.fullmatch(name):
            diagnostics.error(
                "SQL-011",
                f"El bind :{name[:60]} no es un identificador válido de 1 a 30 caracteres.",
            )
        elif name not in binds:
            binds.append(name)
    for match in re.finditer(r":", masked):
        position = match.start()
        previous = masked[position - 1] if position else ""
        following = masked[position + 1] if position + 1 < len(masked) else ""
        if previous == ":" or following == ":" or re.match(r"[A-Za-z]", following):
            continue
        diagnostics.error(
            "SQL-012",
            "Solo se admiten binds con nombre que comience por una letra, por ejemplo :P_CLIENTE.",
            location=f"carácter {position + 1}",
        )
    return tuple(binds)
