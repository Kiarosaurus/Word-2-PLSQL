from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path, PurePosixPath
import re
import zipfile
from xml.etree import ElementTree as ET

from .diagnostics import Diagnostics


MAX_ARCHIVE_BYTES = 10 * 1024 * 1024
MAX_MEMBER_COUNT = 500
MAX_MEMBER_BYTES = 10 * 1024 * 1024
MAX_TOTAL_UNCOMPRESSED = 50 * 1024 * 1024
MAX_COMPRESSION_RATIO = 250

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REQUIRED_PARTS = {
    "[content_types].xml",
    "_rels/.rels",
    "word/document.xml",
}


@dataclass(frozen=True, slots=True)
class ZipInspection:
    xml_parts: tuple[str, ...]


def _unsafe_name(name: str) -> bool:
    if "\x00" in name or "\\" in name or name.startswith(("/", "\\")):
        return True
    path = PurePosixPath(name)
    return path.is_absolute() or any(part in {"", ".."} for part in path.parts)


def _prohibited_part(name: str) -> str | None:
    lowered = name.lower()
    prohibited = {
        "word/vbaproject.bin": "macro VBA",
        "word/comments.xml": "comentarios",
        "word/people.xml": "comentarios modernos",
    }
    if lowered in prohibited:
        return prohibited[lowered]
    prefixes = {
        "customxml/": "XML personalizado",
        "word/activex/": "control ActiveX",
        "word/embeddings/": "objeto incrustado",
        "word/media/": "imagen",
        "word/diagrams/": "SmartArt",
        "word/charts/": "gráfico",
        "word/comments": "comentarios",
    }
    for prefix, label in prefixes.items():
        if lowered.startswith(prefix):
            return label
    return None


def inspect_docx_zip(path: Path, diagnostics: Diagnostics) -> ZipInspection | None:
    location = str(path)
    if path.suffix.lower() != ".docx":
        diagnostics.error(
            "FILE-001",
            "La plantilla debe tener extensión .docx.",
            location=location,
            suggestion="Guarde el documento como Documento de Word (*.docx), sin macros.",
        )
        return None
    if not path.is_file():
        diagnostics.error("FILE-002", "No existe la plantilla DOCX.", location=location)
        return None
    try:
        size = path.stat().st_size
    except OSError as exc:
        diagnostics.error("FILE-003", f"No se pudo leer la plantilla: {exc}", location=location)
        return None
    if size > MAX_ARCHIVE_BYTES:
        diagnostics.error(
            "ZIP-001",
            f"La plantilla supera el límite de {MAX_ARCHIVE_BYTES // 1024 // 1024} MiB.",
            location=location,
        )
        return None

    xml_parts: list[str] = []
    try:
        with zipfile.ZipFile(path, "r") as archive:
            infos = archive.infolist()
            if len(infos) > MAX_MEMBER_COUNT:
                diagnostics.error(
                    "ZIP-002",
                    f"El DOCX contiene demasiadas partes ({len(infos)}).",
                    location=location,
                )
                return None

            seen: set[str] = set()
            total = 0
            for info in infos:
                name = info.filename
                normalized = name.casefold()
                if _unsafe_name(name):
                    diagnostics.error("ZIP-003", "El DOCX contiene una ruta insegura.", location=name)
                    continue
                if normalized in seen:
                    diagnostics.error("ZIP-004", "El DOCX contiene una parte duplicada.", location=name)
                    continue
                seen.add(normalized)
                if info.flag_bits & 0x1:
                    diagnostics.error("ZIP-005", "El DOCX contiene una parte cifrada.", location=name)
                if info.file_size > MAX_MEMBER_BYTES:
                    diagnostics.error("ZIP-006", "Una parte del DOCX supera el tamaño permitido.", location=name)
                total += info.file_size
                if total > MAX_TOTAL_UNCOMPRESSED:
                    diagnostics.error("ZIP-007", "El DOCX excede el tamaño descomprimido permitido.", location=location)
                    break
                if info.file_size and info.compress_size == 0:
                    diagnostics.error("ZIP-008", "Se detectó una relación de compresión inválida.", location=name)
                elif info.compress_size and info.file_size / info.compress_size > MAX_COMPRESSION_RATIO:
                    diagnostics.error("ZIP-009", "Se detectó una posible bomba ZIP.", location=name)

                label = _prohibited_part(name)
                if label:
                    diagnostics.error(
                        "OOXML-001",
                        f"La plantilla contiene un elemento no permitido: {label}.",
                        location=name,
                    )
                if normalized.endswith((".xml", ".rels")):
                    xml_parts.append(name)

            for required in sorted(REQUIRED_PARTS - seen):
                diagnostics.error(
                    "OOXML-006",
                    "Falta una parte obligatoria del documento Word.",
                    location=required,
                )

            if diagnostics.has_errors:
                return None

            for name in xml_parts:
                data = archive.read(name)
                # El DTD puede aparecer después de comentarios o espacios muy
                # largos. Los límites ZIP anteriores acotan el costo de revisar
                # la parte completa y evitan que una declaración tardía eluda
                # esta defensa.
                probe = data.upper()
                if b"<!DOCTYPE" in probe or b"<!ENTITY" in probe:
                    diagnostics.error(
                        "OOXML-002",
                        "No se permiten DTD ni entidades XML.",
                        location=name,
                    )
                    continue
                try:
                    root = ET.fromstring(data)
                except ET.ParseError as exc:
                    diagnostics.error("OOXML-003", f"XML no válido: {exc}", location=name)
                    continue
                if name.lower().endswith(".rels"):
                    for relationship in root.findall(f"{{{REL_NS}}}Relationship"):
                        if relationship.attrib.get("TargetMode", "").casefold() == "external":
                            diagnostics.error(
                                "OOXML-004",
                                "La plantilla contiene una relación externa.",
                                location=name,
                                suggestion="Quite hipervínculos, plantillas adjuntas o recursos externos.",
                            )

            content_types = archive.read("[Content_Types].xml") if "[Content_Types].xml" in archive.namelist() else b""
            if re.search(br"macroEnabled|vbaProject", content_types, re.IGNORECASE):
                diagnostics.error("OOXML-005", "La plantilla declara contenido habilitado para macros.", location="[Content_Types].xml")
    except (OSError, zipfile.BadZipFile, KeyError) as exc:
        diagnostics.error("ZIP-010", f"El archivo no es un DOCX válido: {exc}", location=location)
        return None

    if diagnostics.has_errors:
        return None
    return ZipInspection(tuple(sorted(xml_parts)))
