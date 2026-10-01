from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Iterable, Literal


Severity = Literal["ERROR", "WARNING"]


@dataclass(frozen=True, slots=True)
class Diagnostic:
    code: str
    severity: Severity
    message: str
    location: str | None = None
    suggestion: str | None = None

    def to_dict(self) -> dict[str, str]:
        return {key: value for key, value in asdict(self).items() if value is not None}


class Diagnostics:
    def __init__(self, *, strict: bool = True) -> None:
        self.strict = strict
        self._items: list[Diagnostic] = []

    @property
    def items(self) -> tuple[Diagnostic, ...]:
        return tuple(self._items)

    @property
    def has_errors(self) -> bool:
        return any(item.severity == "ERROR" for item in self._items)

    @property
    def error_count(self) -> int:
        return sum(item.severity == "ERROR" for item in self._items)

    @property
    def warning_count(self) -> int:
        return sum(item.severity == "WARNING" for item in self._items)

    def error(
        self,
        code: str,
        message: str,
        *,
        location: str | None = None,
        suggestion: str | None = None,
    ) -> None:
        self._items.append(Diagnostic(code, "ERROR", message, location, suggestion))

    def warning(
        self,
        code: str,
        message: str,
        *,
        location: str | None = None,
        suggestion: str | None = None,
        strict_error: bool = False,
    ) -> None:
        severity: Severity = "ERROR" if strict_error and self.strict else "WARNING"
        self._items.append(Diagnostic(code, severity, message, location, suggestion))

    def extend(self, values: Iterable[Diagnostic]) -> None:
        self._items.extend(values)

    def to_dict(self) -> dict[str, object]:
        ordered = sorted(
            self._items,
            key=lambda item: (
                0 if item.severity == "ERROR" else 1,
                item.code,
                item.location or "",
                item.message,
            ),
        )
        return {
            "valid": not self.has_errors,
            "errors": self.error_count,
            "warnings": self.warning_count,
            "diagnostics": [item.to_dict() for item in ordered],
        }


class CompilationError(Exception):
    """Error controlado cuya explicación ya está en Diagnostics."""

