"""Headless find-and-replace across Excel workbooks.

Files are treated as pure data: no COM, no Excel process, no GUI.
"""

from __future__ import annotations

import fnmatch
import os
import re
import shutil
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Iterator, Sequence

XLSX_SUFFIXES = {".xlsx", ".xlsm"}
XLS_SUFFIXES = {".xls"}
SUPPORTED_SUFFIXES = XLSX_SUFFIXES | XLS_SUFFIXES

SKIP_PREFIXES = ("~$",)


@dataclass
class Rule:
    """One find/replace pair."""

    find: str
    replace: str
    match_case: bool = True
    whole_cell: bool = False
    regex: bool = False

    def compile(self) -> "CompiledRule":
        flags = 0 if self.match_case else re.IGNORECASE
        pattern = self.find if self.regex else re.escape(self.find)
        if self.whole_cell:
            pattern = rf"\A(?:{pattern})\Z"
        repl = self.replace if self.regex else self.replace.replace("\\", "\\\\")
        return CompiledRule(re.compile(pattern, flags), repl)


@dataclass
class CompiledRule:
    pattern: re.Pattern
    replacement: str

    def apply(self, text: str) -> tuple[str, int]:
        new_text, count = self.pattern.subn(self.replacement, text)
        return new_text, count


@dataclass
class FileResult:
    path: Path
    replacements: int = 0
    cells: int = 0
    sheets_touched: list[str] = field(default_factory=list)
    written: bool = False
    backup: Path | None = None
    error: str | None = None
    skipped_reason: str | None = None

    @property
    def ok(self) -> bool:
        return self.error is None


@dataclass
class RunSummary:
    results: list[FileResult] = field(default_factory=list)

    @property
    def total_replacements(self) -> int:
        return sum(r.replacements for r in self.results)

    @property
    def total_cells(self) -> int:
        return sum(r.cells for r in self.results)

    @property
    def changed_files(self) -> list[FileResult]:
        return [r for r in self.results if r.replacements > 0]

    @property
    def failed_files(self) -> list[FileResult]:
        return [r for r in self.results if r.error]

    @property
    def skipped_files(self) -> list[FileResult]:
        return [r for r in self.results if r.skipped_reason]


def discover_files(
    root: Path,
    recursive: bool = True,
    include: Sequence[str] | None = None,
    exclude: Sequence[str] | None = None,
) -> list[Path]:
    """List supported workbooks under *root*, skipping Excel lock files."""
    root = Path(root)
    if root.is_file():
        candidates: Iterable[Path] = [root]
    else:
        pattern = "**/*" if recursive else "*"
        candidates = sorted(p for p in root.glob(pattern) if p.is_file())

    found: list[Path] = []
    for path in candidates:
        if path.suffix.lower() not in SUPPORTED_SUFFIXES:
            continue
        if path.name.startswith(SKIP_PREFIXES):
            continue
        if include and not any(fnmatch.fnmatch(path.name, pat) for pat in include):
            continue
        if exclude and any(fnmatch.fnmatch(path.name, pat) for pat in exclude):
            continue
        found.append(path)
    return found


def _replace_text(text: str, rules: Sequence[CompiledRule]) -> tuple[str, int]:
    total = 0
    for rule in rules:
        text, count = rule.apply(text)
        total += count
    return text, total


def _backup_path(path: Path) -> Path:
    candidate = path.with_suffix(path.suffix + ".bak")
    counter = 1
    while candidate.exists():
        candidate = path.with_suffix(f"{path.suffix}.bak{counter}")
        counter += 1
    return candidate


def _process_xlsx(
    path: Path,
    rules: Sequence[CompiledRule],
    apply: bool,
    backup: bool,
    include_formulas: bool,
    include_sheet_names: bool,
) -> FileResult:
    import openpyxl

    result = FileResult(path=path)
    workbook = openpyxl.load_workbook(
        path, keep_vba=path.suffix.lower() == ".xlsm", data_only=False, rich_text=False
    )
    try:
        for sheet in workbook.worksheets:
            sheet_hits = 0
            for row in sheet.iter_rows():
                for cell in row:
                    value = cell.value
                    if not isinstance(value, str):
                        continue
                    if value.startswith("=") and not include_formulas:
                        continue
                    new_value, count = _replace_text(value, rules)
                    if count:
                        sheet_hits += count
                        result.cells += 1
                        cell.value = new_value
            if include_sheet_names:
                new_title, count = _replace_text(sheet.title, rules)
                if count and new_title and new_title != sheet.title:
                    sheet.title = new_title
                    sheet_hits += count
            if sheet_hits:
                result.replacements += sheet_hits
                result.sheets_touched.append(sheet.title)

        if result.replacements and apply:
            result.backup = _save_in_place(workbook.save, path, backup)
            result.written = True
    finally:
        workbook.close()
    return result


def _save_in_place(save, path: Path, backup: bool) -> Path | None:
    """Write to a sibling temp file first, so a failure leaves the original intact."""
    tmp = path.with_name(path.name + ".emr-tmp")
    backup_path = None
    try:
        save(tmp)
        if backup:
            backup_path = _backup_path(path)
            shutil.copy2(path, backup_path)
        os.replace(tmp, path)
    except BaseException:
        tmp.unlink(missing_ok=True)
        if backup_path is not None:
            backup_path.unlink(missing_ok=True)
        raise
    return backup_path


def _write_xls_keeping_style(out_sheet, row_index: int, col_index: int, value: str) -> None:
    """xlwt.write() resets a cell's format, so carry the original xf index over."""
    cells = out_sheet.row(row_index)._Row__cells
    original = cells.get(col_index)
    xf_idx = getattr(original, "xf_idx", None)
    out_sheet.write(row_index, col_index, value)
    if xf_idx is not None:
        written = cells.get(col_index)
        if written is not None:
            written.xf_idx = xf_idx


def _process_xls(
    path: Path,
    rules: Sequence[CompiledRule],
    apply: bool,
    backup: bool,
    include_formulas: bool,
    include_sheet_names: bool,
) -> FileResult:
    import xlrd
    from xlutils.copy import copy as xl_copy

    result = FileResult(path=path)
    book = xlrd.open_workbook(path, formatting_info=True, on_demand=False)
    out_book = None

    for sheet_index in range(book.nsheets):
        sheet = book.sheet_by_index(sheet_index)
        sheet_hits = 0
        pending: list[tuple[int, int, str]] = []
        for row_index in range(sheet.nrows):
            for col_index in range(sheet.ncols):
                cell = sheet.cell(row_index, col_index)
                if cell.ctype != xlrd.XL_CELL_TEXT:
                    continue
                new_value, count = _replace_text(cell.value, rules)
                if count:
                    sheet_hits += count
                    result.cells += 1
                    pending.append((row_index, col_index, new_value))
        new_title = sheet.name
        title_hits = 0
        if include_sheet_names:
            new_title, title_hits = _replace_text(sheet.name, rules)
            if not new_title:
                new_title, title_hits = sheet.name, 0

        if pending or title_hits:
            result.replacements += sheet_hits + title_hits
            result.sheets_touched.append(sheet.name)
            if apply:
                if out_book is None:
                    out_book = xl_copy(book)
                out_sheet = out_book.get_sheet(sheet_index)
                for row_index, col_index, value in pending:
                    _write_xls_keeping_style(out_sheet, row_index, col_index, value)
                if title_hits:
                    out_sheet.set_name(new_title)

    if out_book is not None:
        result.backup = _save_in_place(
            lambda target: out_book.save(str(target)), path, backup
        )
        result.written = True

    if include_formulas:
        result.skipped_reason = "formula text is not editable in legacy .xls"
    return result


def process_file(
    path: Path,
    rules: Sequence[Rule],
    apply: bool = False,
    backup: bool = True,
    include_formulas: bool = False,
    include_sheet_names: bool = False,
) -> FileResult:
    """Run every rule over one workbook. Never raises: errors land in the result."""
    path = Path(path)
    compiled = [rule.compile() for rule in rules]
    suffix = path.suffix.lower()
    try:
        if suffix in XLSX_SUFFIXES:
            return _process_xlsx(
                path, compiled, apply, backup, include_formulas, include_sheet_names
            )
        if suffix in XLS_SUFFIXES:
            return _process_xls(
                path, compiled, apply, backup, include_formulas, include_sheet_names
            )
        return FileResult(path=path, skipped_reason=f"unsupported file type {suffix}")
    except Exception as exc:  # one bad workbook must not stop the batch
        return FileResult(path=path, error=f"{type(exc).__name__}: {exc}")


def _worker(args) -> FileResult:
    path, rules, apply, backup, include_formulas, include_sheet_names = args
    return process_file(path, rules, apply, backup, include_formulas, include_sheet_names)


def run(
    files: Sequence[Path],
    rules: Sequence[Rule],
    apply: bool = False,
    backup: bool = True,
    include_formulas: bool = False,
    include_sheet_names: bool = False,
    workers: int | None = None,
    progress: Callable[[FileResult], None] | None = None,
) -> RunSummary:
    """Process every file, in parallel when there is enough work to justify it."""
    summary = RunSummary()
    if not files:
        return summary

    if workers is None:
        workers = min(len(files), (os.cpu_count() or 2))
    workers = max(1, min(workers, len(files)))

    payload = [
        (Path(f), list(rules), apply, backup, include_formulas, include_sheet_names)
        for f in files
    ]

    if workers == 1:
        for args in payload:
            result = _worker(args)
            summary.results.append(result)
            if progress:
                progress(result)
        return summary

    with ProcessPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(_worker, args): args[0] for args in payload}
        for future in as_completed(futures):
            try:
                result = future.result()
            except Exception as exc:
                result = FileResult(
                    path=futures[future], error=f"{type(exc).__name__}: {exc}"
                )
            summary.results.append(result)
            if progress:
                progress(result)
    summary.results.sort(key=lambda r: str(r.path))
    return summary
