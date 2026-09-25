"""Excel Mass Replacer — headless find-and-replace across Excel files."""

from .core import (
    FileResult,
    Rule,
    RunSummary,
    discover_files,
    process_file,
    run,
)

__version__ = "1.0.0"
__all__ = [
    "FileResult",
    "Rule",
    "RunSummary",
    "discover_files",
    "process_file",
    "run",
    "__version__",
]
