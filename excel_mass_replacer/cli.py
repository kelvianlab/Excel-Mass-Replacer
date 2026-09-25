"""Command-line interface for Excel Mass Replacer."""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

from . import __version__
from .core import Rule, RunSummary, discover_files, run


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-mass-replacer",
        description=(
            "Find and replace text across every Excel file in a folder, "
            "without opening Microsoft Excel."
        ),
        epilog=(
            "By default nothing is written: you get a preview (dry run). "
            "Add --apply to actually change the files."
        ),
    )
    parser.add_argument(
        "find", nargs="?", help="text to look for (omit when using --pairs-file)"
    )
    parser.add_argument("replace", nargs="?", help="text to put in its place")
    parser.add_argument(
        "-d",
        "--dir",
        default=".",
        help="folder (or single file) to process; default: current folder",
    )
    parser.add_argument(
        "--pairs-file",
        help="tab-separated file of find<TAB>replace lines, one rule per line",
    )
    parser.add_argument(
        "--apply",
        action="store_true",
        help="write the changes; without it the run is a preview only",
    )
    parser.add_argument(
        "--no-backup",
        action="store_true",
        help="do not keep a .bak copy of each changed file (only with --apply)",
    )
    parser.add_argument(
        "--no-recursive",
        action="store_true",
        help="only the given folder, do not descend into subfolders",
    )
    parser.add_argument(
        "-i",
        "--ignore-case",
        action="store_true",
        help="match regardless of upper/lower case",
    )
    parser.add_argument(
        "-w",
        "--whole-cell",
        action="store_true",
        help="only replace when the whole cell equals the search text",
    )
    parser.add_argument(
        "--regex", action="store_true", help="treat the search text as a regular expression"
    )
    parser.add_argument(
        "--include-formulas",
        action="store_true",
        help="also rewrite text inside .xlsx formulas (ignored for legacy .xls)",
    )
    parser.add_argument(
        "--sheet-names",
        action="store_true",
        help="also rename worksheet tabs that match",
    )
    parser.add_argument(
        "--include",
        action="append",
        metavar="GLOB",
        help="only files whose name matches this pattern (repeatable)",
    )
    parser.add_argument(
        "--exclude",
        action="append",
        metavar="GLOB",
        help="skip files whose name matches this pattern (repeatable)",
    )
    parser.add_argument(
        "-j",
        "--workers",
        type=int,
        default=None,
        help="number of parallel workers; default: one per CPU core",
    )
    parser.add_argument("-q", "--quiet", action="store_true", help="only print the summary")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def load_rules(args: argparse.Namespace) -> list[Rule]:
    def make(find: str, replace: str) -> Rule:
        return Rule(
            find=find,
            replace=replace,
            match_case=not args.ignore_case,
            whole_cell=args.whole_cell,
            regex=args.regex,
        )

    if args.pairs_file:
        rules: list[Rule] = []
        path = Path(args.pairs_file)
        for line_no, raw in enumerate(
            path.read_text(encoding="utf-8-sig").splitlines(), start=1
        ):
            if not raw.strip() or raw.lstrip().startswith("#"):
                continue
            if "\t" not in raw:
                raise ValueError(
                    f"{path}:{line_no}: expected 'find<TAB>replace', got: {raw!r}"
                )
            find, replace = raw.split("\t", 1)
            if not find:
                raise ValueError(f"{path}:{line_no}: search text must not be empty")
            rules.append(make(find, replace))
        if not rules:
            raise ValueError(f"{path}: no rules found")
        return rules

    if args.find is None or args.replace is None:
        raise ValueError("give both FIND and REPLACE, or use --pairs-file")
    if not args.find:
        raise ValueError("search text must not be empty")
    return [make(args.find, args.replace)]


def print_report(summary: RunSummary, applied: bool, quiet: bool) -> None:
    if not quiet:
        for result in summary.results:
            if result.error:
                print(f"  ERROR  {result.path}: {result.error}")
            elif result.replacements:
                mark = "CHANGED" if result.written else "  WOULD "
                sheets = ", ".join(result.sheets_touched)
                print(
                    f"  {mark} {result.path}: {result.replacements} replacement(s) "
                    f"in {result.cells} cell(s) [{sheets}]"
                )

    print()
    print(f"Files scanned     : {len(summary.results)}")
    print(f"Files with matches: {len(summary.changed_files)}")
    print(f"Cells matched     : {summary.total_cells}")
    print(f"Replacements      : {summary.total_replacements}")
    if summary.failed_files:
        print(f"Files failed      : {len(summary.failed_files)}")
    if not applied:
        print()
        print("Preview only — nothing was written. Re-run with --apply to save changes.")


def main(argv: list[str] | None = None) -> int:
    import multiprocessing

    multiprocessing.freeze_support()
    args = build_parser().parse_args(argv)

    try:
        rules = load_rules(args)
    except (ValueError, OSError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    root = Path(args.dir).expanduser()
    if not root.exists():
        print(f"error: path not found: {root}", file=sys.stderr)
        return 2

    files = discover_files(
        root,
        recursive=not args.no_recursive,
        include=args.include,
        exclude=args.exclude,
    )
    if not files:
        print(f"No .xlsx/.xlsm/.xls files found in {root.resolve()}")
        return 0

    print(f"Scanning {len(files)} file(s) in {root.resolve()}")
    if args.apply:
        print("Mode: APPLY (files will be modified)")
    else:
        print("Mode: PREVIEW (dry run, nothing will be modified)")
    print()

    started = time.perf_counter()
    summary = run(
        files,
        rules,
        apply=args.apply,
        backup=not args.no_backup,
        include_formulas=args.include_formulas,
        include_sheet_names=args.sheet_names,
        workers=args.workers,
    )
    elapsed = time.perf_counter() - started

    print_report(summary, applied=args.apply, quiet=args.quiet)
    print(f"Finished in {elapsed:.2f}s")

    return 1 if summary.failed_files else 0


if __name__ == "__main__":
    raise SystemExit(main())
