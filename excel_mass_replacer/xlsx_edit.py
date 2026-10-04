"""Text replacement performed directly on the parts of an .xlsx/.xlsm package.

A workbook is a ZIP of XML parts. Loading one through a workbook library and
saving it again re-serialises every part, so whatever that library does not
model is dropped on the way out: cached formula results, images, form
controls, printer settings, threaded comments. Here only the text that
actually changes is rewritten, and every other byte is copied through, so
anything this module does not understand survives untouched.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Sequence

SHARED_STRINGS = "xl/sharedStrings.xml"
WORKBOOK = "xl/workbook.xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"


@dataclass
class EditReport:
    replacements: int = 0
    cells: int = 0
    sheets_touched: list[str] = field(default_factory=list)
    dropdown_lists: int = 0
    sheet_names: int = 0

    @property
    def changed(self) -> bool:
        return self.replacements > 0


_ENTITY = re.compile(r"&(?:#x([0-9A-Fa-f]+)|#(\d+)|(amp|lt|gt|quot|apos));")
_NAMED = {"amp": "&", "lt": "<", "gt": ">", "quot": '"', "apos": "'"}


def _unescape(text: str) -> str:
    def one(match: re.Match) -> str:
        hexa, dec, name = match.groups()
        if name:
            return _NAMED[name]
        return chr(int(hexa, 16) if hexa else int(dec))

    return _ENTITY.sub(one, text)


def _escape(text: str) -> str:
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return text.replace("\r", "&#13;").replace("\n", "&#10;").replace("\t", "&#9;")


def _escape_attr(text: str) -> str:
    return _escape(text).replace('"', "&quot;")


def _apply(text: str, rules: Sequence) -> tuple[str, int]:
    total = 0
    for rule in rules:
        text, count = rule.apply(text)
        total += count
    return text, total


def _sub_raw(raw: str, rules: Sequence) -> tuple[str, int]:
    """Replace inside one escaped XML text node, leaving it alone if nothing hits."""
    new_text, count = _apply(_unescape(raw), rules)
    if not count:
        return raw, 0
    return _escape(new_text), count


_T_ELEM = re.compile(r"(<t(?:\s[^>]*)?>)(.*?)(</t>)", re.DOTALL)
_SI_ELEM = re.compile(r"(<si(?:\s[^>]*)?>)(.*?)(</si>)", re.DOTALL)
_RPH = re.compile(r"<rPh\b.*?</rPh>", re.DOTALL)
_IS_ELEM = re.compile(r"(<is(?:\s[^>]*)?>)(.*?)(</is>)", re.DOTALL)
_F_ELEM = re.compile(r"(<(?:xm:)?f(?:\s[^>]*)?>)(.*?)(</(?:xm:)?f>)", re.DOTALL)
_DV_BLOCK = re.compile(
    r"<(?:x14:)?dataValidation\b[^>]*>.*?</(?:x14:)?dataValidation>", re.DOTALL
)
_DV_LIST_FORMULA = re.compile(
    r"(<(?:x14:)?formula1(?:\s[^>]*)?>(?:<xm:f>)?)(.*?)((?:</xm:f>)?</(?:x14:)?formula1>)",
    re.DOTALL,
)
_SST_CELL = re.compile(r'<c\b(?=[^>]*\bt="s")[^>]*>\s*<v>(\d+)</v>', re.DOTALL)
_SHEET_TAG = re.compile(r'<sheet\b[^>]*\bname="([^"]*)"[^>]*>')
_REL = re.compile(r'<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*>')


def _splice(xml: str, edits: list[tuple[int, int, str]]) -> str:
    """Apply (start, end, replacement) edits, which must already be in order."""
    if not edits:
        return xml
    out: list[str] = []
    last = 0
    for start, end, text in edits:
        out.append(xml[last:start])
        out.append(text)
        last = end
    out.append(xml[last:])
    return "".join(out)


def _rewrite_one_string(inner: str, rules: Sequence) -> tuple[str, int]:
    """Rewrite one shared string, which may be split into several styled runs."""
    skip = [(m.start(), m.end()) for m in _RPH.finditer(inner)]
    runs = [
        m
        for m in _T_ELEM.finditer(inner)
        if not any(start <= m.start() < end for start, end in skip)
    ]
    if not runs:
        return inner, 0

    # Run by run first, so each run keeps its own formatting.
    edits: list[tuple[int, int, str]] = []
    total = 0
    for match in runs:
        new_raw, count = _sub_raw(match.group(2), rules)
        if count:
            edits.append((match.start(2), match.end(2), new_raw))
            total += count
    if total:
        return _splice(inner, edits), total

    # Nothing matched run by run: the search text may straddle two runs, or a
    # whole-cell rule may need the entire string. Both only work on the joined
    # text, and the result has to collapse into the first run.
    joined = "".join(_unescape(m.group(2)) for m in runs)
    replaced, count = _apply(joined, rules)
    if not count:
        return inner, 0
    edits = [
        (m.start(2), m.end(2), _escape(replaced) if i == 0 else "")
        for i, m in enumerate(runs)
    ]
    return _splice(inner, edits), count


def _rewrite_shared_strings(xml: str, rules: Sequence) -> tuple[str, dict[int, int]]:
    """Returns the new XML and {shared string index: replacements in it}."""
    hits: dict[int, int] = {}
    edits: list[tuple[int, int, str]] = []
    for index, match in enumerate(_SI_ELEM.finditer(xml)):
        new_inner, count = _rewrite_one_string(match.group(2), rules)
        if count:
            hits[index] = count
            edits.append((match.start(2), match.end(2), new_inner))
    return _splice(xml, edits), hits


def _rewrite_dropdowns(xml: str, rules: Sequence) -> tuple[str, int, int]:
    """Replace inside dropdown lists that are typed into the rule itself.

    A list that points at cells (``$A$1:$A$9``, ``Supplier!$A$2:$A$371``) needs
    nothing here: its text lives in those cells and is replaced with the rest.
    """
    edits: list[tuple[int, int, str]] = []
    total = 0
    lists = 0
    for block in _DV_BLOCK.finditer(xml):
        if 'type="list"' not in block.group(0):
            continue
        for formula in _DV_LIST_FORMULA.finditer(block.group(0)):
            raw = formula.group(2)
            if not _unescape(raw).lstrip().startswith('"'):
                continue  # a reference, not a typed-in list
            new_raw, count = _sub_raw(raw, rules)
            if count:
                start = block.start() + formula.start(2)
                end = block.start() + formula.end(2)
                edits.append((start, end, new_raw))
                total += count
                lists += 1
    return _splice(xml, edits), total, lists


def _rewrite_sheet(
    xml: str, rules: Sequence, sst_hits: dict[int, int], include_formulas: bool
) -> tuple[str, int, int, int]:
    """Returns new XML, replacements, cells changed, dropdown lists changed."""
    replacements = 0
    cells = 0

    # Cells that point at a shared string this run already rewrote.
    for match in _SST_CELL.finditer(xml):
        count = sst_hits.get(int(match.group(1)))
        if count:
            replacements += count
            cells += 1

    edits: list[tuple[int, int, str]] = []

    for block in _IS_ELEM.finditer(xml):
        new_inner, count = _rewrite_one_string(block.group(2), rules)
        if count:
            edits.append((block.start(2), block.end(2), new_inner))
            replacements += count
            cells += 1

    if include_formulas:
        for match in _F_ELEM.finditer(xml):
            new_raw, count = _sub_raw(match.group(2), rules)
            if count:
                edits.append((match.start(2), match.end(2), new_raw))
                replacements += count
                cells += 1

    xml = _splice(xml, sorted(edits))
    xml, dropdown_hits, dropdown_lists = _rewrite_dropdowns(xml, rules)
    replacements += dropdown_hits
    return xml, replacements, cells, dropdown_lists


def _sheet_names(parts: dict[str, bytes]) -> dict[str, str]:
    """Map worksheet part name to the sheet name shown on its tab."""
    workbook = parts.get(WORKBOOK, b"").decode("utf-8", "replace")
    rels = parts.get(WORKBOOK_RELS, b"").decode("utf-8", "replace")
    targets = {m.group(1): m.group(2) for m in _REL.finditer(rels)}
    names: dict[str, str] = {}
    for match in _SHEET_TAG.finditer(workbook):
        rid = re.search(r'r:id="([^"]+)"', match.group(0))
        if not rid:
            continue
        target = targets.get(rid.group(1), "")
        target = target.split("/")[-1]
        if target:
            names[f"xl/worksheets/{target}"] = _unescape(match.group(1))
    return names


def _rename_sheets(
    parts: dict[str, bytes], rules: Sequence
) -> tuple[dict[str, str], int]:
    """Rename sheet tabs in workbook.xml. Returns {old: new} and the hit count."""
    xml = parts[WORKBOOK].decode("utf-8")
    renames: dict[str, str] = {}
    edits: list[tuple[int, int, str]] = []
    total = 0
    for match in _SHEET_TAG.finditer(xml):
        old = _unescape(match.group(1))
        new, count = _apply(old, rules)
        if not count or not new or new == old:
            continue
        renames[old] = new
        total += count
        start = match.start(1)
        edits.append((start, match.end(1), _escape_attr(new)))
    if edits:
        parts[WORKBOOK] = _splice(xml, edits).encode("utf-8")
    return renames, total


def _retarget_formulas(parts: dict[str, bytes], renames: dict[str, str]) -> None:
    """Point formulas and defined names at the new sheet names after a rename."""
    if not renames:
        return
    pairs = []
    for old, new in renames.items():
        pairs.append((f"'{_escape(old)}'!", f"'{_escape(new)}'!"))
        pairs.append((f"{_escape(old)}!", f"{_escape(new)}!"))

    def fix(xml: str) -> str:
        def one(match: re.Match) -> str:
            body = match.group(2)
            for old, new in pairs:
                body = body.replace(old, new)
            return match.group(1) + body + match.group(3)

        xml = _F_ELEM.sub(one, xml)
        return re.sub(
            r"(<definedName\b[^>]*>)(.*?)(</definedName>)", one, xml, flags=re.DOTALL
        )

    for name in list(parts):
        if name.startswith("xl/worksheets/") or name == WORKBOOK:
            parts[name] = fix(parts[name].decode("utf-8")).encode("utf-8")


def rewrite(
    src: Path,
    rules: Sequence,
    include_formulas: bool = False,
    include_sheet_names: bool = False,
) -> tuple[list[tuple[zipfile.ZipInfo, bytes]], EditReport]:
    """Read a workbook and return its parts with every rule applied.

    Nothing is written here. Parts that need no change keep their original
    bytes, so saving these back out leaves the untouched ones identical.
    """
    report = EditReport()
    order: list[zipfile.ZipInfo] = []
    parts: dict[str, bytes] = {}
    with zipfile.ZipFile(src) as archive:
        for info in archive.infolist():
            order.append(info)
            parts[info.filename] = archive.read(info.filename)

    tab_names = _sheet_names(parts)

    sst_hits: dict[int, int] = {}
    if SHARED_STRINGS in parts:
        xml = parts[SHARED_STRINGS].decode("utf-8")
        new_xml, sst_hits = _rewrite_shared_strings(xml, rules)
        if sst_hits:
            parts[SHARED_STRINGS] = new_xml.encode("utf-8")

    for name in sorted(parts):
        if not name.startswith("xl/worksheets/") or not name.endswith(".xml"):
            continue
        xml = parts[name].decode("utf-8")
        new_xml, replacements, cells, lists = _rewrite_sheet(
            xml, rules, sst_hits, include_formulas
        )
        if new_xml != xml:
            parts[name] = new_xml.encode("utf-8")
        if replacements:
            report.replacements += replacements
            report.cells += cells
            report.dropdown_lists += lists
            report.sheets_touched.append(tab_names.get(name, name))

    if include_sheet_names and WORKBOOK in parts:
        renames, hits = _rename_sheets(parts, rules)
        if hits:
            _retarget_formulas(parts, renames)
            report.replacements += hits
            report.sheet_names += len(renames)
            for new in renames.values():
                if new not in report.sheets_touched:
                    report.sheets_touched.append(new)

    return [(info, parts[info.filename]) for info in order], report


def write_package(parts: Sequence[tuple[zipfile.ZipInfo, bytes]], dest: Path) -> None:
    """Write the parts back out, keeping each one's original ZIP metadata."""
    with zipfile.ZipFile(dest, "w") as archive:
        for info, data in parts:
            archive.writestr(info, data)
