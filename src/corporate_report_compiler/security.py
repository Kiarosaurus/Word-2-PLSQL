from __future__ import annotations

from dataclasses import dataclass
import posixpath
from pathlib import Path, PurePosixPath
import re
import zipfile
import zlib
from xml.etree import ElementTree as ET
from xml.parsers import expat

from .diagnostics import Diagnostics


MAX_ARCHIVE_BYTES = 10 * 1024 * 1024
MAX_MEMBER_COUNT = 500
MAX_MEMBER_BYTES = 10 * 1024 * 1024
MAX_TOTAL_UNCOMPRESSED = 50 * 1024 * 1024
MAX_COMPRESSION_RATIO = 250

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
OFFICE_DOCUMENT_TYPE = "officedocument"
REQUIRED_PARTS = {
    "[content_types].xml",
    "_rels/.rels",
    "word/document.xml",
}
SUPPORTED_COMPRESSION = {zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED}

# Tipos de relación (último segmento del URI, en minúsculas) que introducen
# contenido activo, binario o ajeno al contrato aunque la parte se renombre.
PROHIBITED_RELATIONSHIP_TYPES = {
    "image": "imagen",
    "oleobject": "objeto incrustado",
    "package": "objeto incrustado",
    "control": "control ActiveX",
    "activexcontrol": "control ActiveX",
    "vbaproject": "macro VBA",
    "comments": "comentarios",
    "commentsextended": "comentarios",
    "commentsids": "comentarios",
    "commentsextensible": "comentarios",
    "chart": "gráfico",
    "diagramdata": "SmartArt",
    "customxml": "XML personalizado",
    "afchunk": "contenido externo insertado",
    "subdocument": "subdocumento",
}


@dataclass(frozen=True, slots=True)
class ZipInspection:
    xml_parts: tuple[str, ...]


def _unsafe_name(name: str) -> bool:
    # ``ZipInfo.filename`` sustituye ``os.sep`` por ``/`` en Windows; el
    # llamador debe pasar también el nombre crudo (``orig_filename``) para que
    # una barra invertida se rechace igual en todas las plataformas.
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


class _DoctypeFound(Exception):
    pass


def _has_doctype(data: bytes) -> bool:
    """Detecta DTD o entidades con el parser, sea cual sea la codificación.

    La búsqueda por bytes ASCII no ve una declaración en UTF-16. Expat
    decodifica la parte según su BOM o declaración y avisa de cualquier
    ``<!DOCTYPE`` o ``<!ENTITY`` antes de procesar el resto del documento.
    """

    probe = data.upper()
    if b"<!DOCTYPE" in probe or b"<!ENTITY" in probe:
        return True

    def mark(*_args: object) -> None:
        raise _DoctypeFound

    parser = expat.ParserCreate()
    parser.StartDoctypeDeclHandler = mark
    parser.EntityDeclHandler = mark
    try:
        parser.Parse(data, True)
    except _DoctypeFound:
        return True
    except (expat.ExpatError, ValueError, LookupError):
        return False
    return False


def _content_type_map(data: bytes) -> tuple[dict[str, str], dict[str, str]]:
    root = ET.fromstring(data)
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    for child in root:
        if child.tag == f"{{{CT_NS}}}Default":
            defaults[child.attrib.get("Extension", "").casefold()] = child.attrib.get("ContentType", "")
        elif child.tag == f"{{{CT_NS}}}Override":
            overrides[child.attrib.get("PartName", "").lstrip("/").casefold()] = child.attrib.get("ContentType", "")
    return defaults, overrides


def _is_xml_part(name: str, defaults: dict[str, str], overrides: dict[str, str]) -> bool:
    lowered = name.casefold()
    if lowered.endswith((".xml", ".rels")):
        return True
    content_type = overrides.get(lowered)
    if content_type is None:
        content_type = defaults.get(PurePosixPath(lowered).suffix.lstrip("."), "")
    return content_type.casefold().endswith(("/xml", "+xml"))


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
                if _unsafe_name(info.orig_filename) or _unsafe_name(name):
                    diagnostics.error("ZIP-003", "El DOCX contiene una ruta insegura.", location=name)
                    continue
                if normalized in seen:
                    diagnostics.error("ZIP-004", "El DOCX contiene una parte duplicada.", location=name)
                    continue
                seen.add(normalized)
                if info.flag_bits & 0x1:
                    diagnostics.error("ZIP-005", "El DOCX contiene una parte cifrada.", location=name)
                if info.compress_type not in SUPPORTED_COMPRESSION:
                    diagnostics.error(
                        "ZIP-011",
                        "El DOCX usa un método de compresión no admitido.",
                        location=name,
                    )
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
                if not normalized.endswith("/"):
                    xml_parts.append(name)

            for required in sorted(REQUIRED_PARTS - seen):
                diagnostics.error(
                    "OOXML-006",
                    "Falta una parte obligatoria del documento Word.",
                    location=required,
                )

            if diagnostics.has_errors:
                return None

            names_by_key = {name.casefold(): name for name in archive.namelist()}
            content_types = archive.read(names_by_key["[content_types].xml"])
            if _has_doctype(content_types):
                diagnostics.error("OOXML-002", "No se permiten DTD ni entidades XML.", location="[Content_Types].xml")
                return None
            try:
                defaults, overrides = _content_type_map(content_types)
            except (ET.ParseError, ValueError, LookupError) as exc:
                diagnostics.error("OOXML-003", f"XML no válido: {exc}", location="[Content_Types].xml")
                return None

            # Toda parte que el paquete declara como XML se inspecciona, aunque
            # su nombre no termine en .xml: python-docx localiza las partes por
            # relación y content type, no por extensión.
            xml_parts = [name for name in xml_parts if _is_xml_part(name, defaults, overrides)]
            for name in xml_parts:
                data = archive.read(name)
                # El DTD puede aparecer después de comentarios o espacios muy
                # largos, o en UTF-16. Los límites ZIP anteriores acotan el
                # costo de revisar la parte completa con el parser.
                if _has_doctype(data):
                    diagnostics.error(
                        "OOXML-002",
                        "No se permiten DTD ni entidades XML.",
                        location=name,
                    )
                    continue
                try:
                    root = ET.fromstring(data)
                except (ET.ParseError, ValueError, LookupError) as exc:
                    diagnostics.error("OOXML-003", f"XML no válido: {exc}", location=name)
                    continue
                if name.lower().endswith(".rels"):
                    for relationship in root.findall(f"{{{REL_NS}}}Relationship"):
                        relation_type = relationship.attrib.get("Type", "").rstrip("/").rsplit("/", 1)[-1].casefold()
                        label = PROHIBITED_RELATIONSHIP_TYPES.get(relation_type)
                        if label:
                            diagnostics.error(
                                "OOXML-001",
                                f"La plantilla contiene un elemento no permitido: {label}.",
                                location=name,
                            )
                        if (
                            name.casefold() == "_rels/.rels"
                            and relation_type == OFFICE_DOCUMENT_TYPE
                            and posixpath.normpath(relationship.attrib.get("Target", "").lstrip("/")).casefold()
                            != "word/document.xml"
                        ):
                            diagnostics.error(
                                "OOXML-007",
                                "La parte principal del documento debe ser word/document.xml.",
                                location=name,
                            )
                        if relationship.attrib.get("TargetMode", "").strip().casefold() == "external":
                            diagnostics.error(
                                "OOXML-004",
                                "La plantilla contiene una relación externa.",
                                location=name,
                                suggestion="Quite hipervínculos, plantillas adjuntas o recursos externos.",
                            )

            if re.search(br"macroEnabled|vbaProject", content_types, re.IGNORECASE):
                diagnostics.error("OOXML-005", "La plantilla declara contenido habilitado para macros.", location="[Content_Types].xml")
    except (
        OSError,
        zipfile.BadZipFile,
        KeyError,
        zlib.error,
        NotImplementedError,
        EOFError,
        RuntimeError,
        ValueError,
    ) as exc:
        diagnostics.error("ZIP-010", f"El archivo no es un DOCX válido: {exc}", location=location)
        return None

    if diagnostics.has_errors:
        return None
    return ZipInspection(tuple(sorted(xml_parts)))
