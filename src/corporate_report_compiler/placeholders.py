from __future__ import annotations

import re

from .diagnostics import Diagnostics


IDENTIFIER_PATTERN = re.compile(r"^[A-Z][A-Z0-9_$#]{0,29}$")
ANY_PLACEHOLDER_PATTERN = re.compile(r"\{\{(.*?)\}\}", re.DOTALL)
COLUMN_PATTERN = re.compile(r"^\{\{\s*COLUMN\s*:\s*([A-Z][A-Z0-9_$#]{0,29})\s*\}\}$", re.IGNORECASE)
FIELD_PATTERN = re.compile(r"\{\{\s*FIELD\s*:\s*([A-Z][A-Z0-9_$#]{0,29})\s*\}\}", re.IGNORECASE)
BUILTINS = {"REPORT_TITLE", "APP_USER", "GENERATED_AT"}


def normalize_identifier(value: str, *, label: str, diagnostics: Diagnostics, location: str | None = None) -> str | None:
    normalized = value.strip().upper()
    if not IDENTIFIER_PATTERN.fullmatch(normalized):
        diagnostics.error(
            "TOKEN-001",
            f"{label} no es un identificador válido: {value!r}.",
            location=location,
            suggestion="Use de 1 a 30 caracteres: letras, números, _, $ o #; empiece con una letra.",
        )
        return None
    return normalized


def validate_text_placeholders(
    text: str,
    *,
    allowed_kinds: set[str],
    diagnostics: Diagnostics,
    location: str,
) -> tuple[str, ...]:
    fields: list[str] = []
    for match in ANY_PLACEHOLDER_PATTERN.finditer(text):
        raw = match.group(0)
        # El package solo recorta espacios ASCII (TRIM); un NBSP o tabulador
        # dentro de las llaves debe rechazarse aquí igual que en ejecución.
        inner = match.group(1).strip(" ")
        upper = inner.upper()
        if upper in BUILTINS:
            if upper not in allowed_kinds:
                diagnostics.error("TOKEN-002", f"El marcador {raw} no está permitido aquí.", location=location)
            continue
        field_match = re.fullmatch(r"FIELD *: *([A-Z][A-Z0-9_$#]{0,29})", upper)
        if field_match:
            if "FIELD" not in allowed_kinds:
                diagnostics.error("TOKEN-002", f"El marcador {raw} no está permitido aquí.", location=location)
            elif field_match.group(1).upper() in BUILTINS:
                diagnostics.error(
                    "TOKEN-005",
                    f"{raw} usa un nombre reservado; escriba {{{{{field_match.group(1).upper()}}}}} sin FIELD:.",
                    location=location,
                )
            else:
                fields.append(field_match.group(1).upper())
            continue
        diagnostics.error(
            "TOKEN-003",
            f"Marcador desconocido o mal formado: {raw}.",
            location=location,
        )
    if text.count("{{") != len(ANY_PLACEHOLDER_PATTERN.findall(text)) or text.count("}}") != len(ANY_PLACEHOLDER_PATTERN.findall(text)):
        diagnostics.error("TOKEN-004", "Hay llaves de marcador sin cerrar o sobrantes.", location=location)
    return tuple(fields)


def parse_column_placeholder(text: str, diagnostics: Diagnostics, location: str) -> str | None:
    match = COLUMN_PATTERN.fullmatch(text.strip())
    if not match:
        diagnostics.error(
            "DOCX-TABLE-004",
            "La celda prototipo debe contener solamente {{COLUMN:ALIAS}}.",
            location=location,
        )
        return None
    return match.group(1).upper()

