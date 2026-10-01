"""API pública del compilador restringido de reportes Oracle APEX."""

__version__ = "1.0.0"

from .compiler import CompilationResult, compile_project, validate_project

__all__ = [
    "CompilationResult",
    "compile_project",
    "validate_project",
    "__version__",
]
