from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass(frozen=True, slots=True)
class TextStyle:
    font_family: str
    font_size: float
    font_weight: str
    font_color: str
    alignment: str | None = None
    background_color: str | None = None

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "font_family": self.font_family,
            "font_size": self.font_size,
            "font_weight": self.font_weight,
            "font_color": self.font_color,
        }
        if self.alignment is not None:
            result["alignment"] = self.alignment
        if self.background_color is not None:
            result["background_color"] = self.background_color
        return result


@dataclass(frozen=True, slots=True)
class ColumnModel:
    name: str
    heading: str
    alignment: str | None
    format_mask: str | None
    inferred_weight: float


@dataclass(frozen=True, slots=True)
class TemplateModel:
    source_path: Path
    source_sha256: str
    header_template: str
    footer_template: str
    fields: tuple[str, ...]
    columns: tuple[ColumnModel, ...]
    template_orientation: str
    title_style: TextStyle
    table_header_style: TextStyle
    table_body_style: TextStyle
    border_width: float
    border_color: str
    footer_style: TextStyle


@dataclass(frozen=True, slots=True)
class ProjectModel:
    source_path: Path
    raw: dict[str, Any]
    report_id: str
    template_path: Path
    query_path: Path
    title: str
    file_name: str
    max_rows: int
    orientation: str | None
    format_item: str | None = None
    orientation_item: str | None = None
    style_overrides: dict[str, Any] = field(default_factory=dict)
    bindings: tuple[dict[str, Any], ...] = field(default_factory=tuple)
    fields: tuple[dict[str, Any], ...] = field(default_factory=tuple)
    excluded_columns: tuple[str, ...] = field(default_factory=tuple)
    column_widths: tuple[dict[str, Any], ...] = field(default_factory=tuple)
    column_overrides: tuple[dict[str, Any], ...] = field(default_factory=tuple)
