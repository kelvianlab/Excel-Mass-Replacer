# Excel Mass Replacer

A lightning-fast, headless Python utility to mass find-and-replace text across all
Excel files (`.xlsx`, `.xlsm`, `.xls`) in a directory — without opening Microsoft Excel.

Excel VBA macros drive a real Excel process through COM. That is why they pop
windows open, hang on a dialog nobody can see, eat memory, and die halfway through a
folder. This tool never touches Excel at all: it reads and rewrites the workbooks as
pure data, so the run is silent, repeatable, and finishes in seconds.

**Measured on this project's own benchmark:** 50 workbooks, 300,000 matching cells,
replaced and saved in **5.9 seconds** on a 4-core machine.

---

## What it does

- Replaces a word or sentence with another one in **every** Excel file in a folder
  (and its subfolders, unless you say otherwise).
- Handles `.xlsx`, `.xlsm` (macro-enabled, macros kept) and legacy `.xls`.
- **Previews by default.** Nothing is written until you explicitly ask for it.
- Keeps a `.bak` copy of every file it changes, unless you turn that off.
- Skips Excel's `~$` lock files, and keeps going when one workbook is corrupt
  instead of aborting the whole batch.
- Runs every CPU core in parallel.
- Never launches Excel, never shows a dialog, never waits for a click.

## What it does not do

- It does not edit Word, PowerPoint, CSV, or ODS files.
- It does not change numbers, dates, or formula *results* — only text.
- It does not undo a run for you. That is what the `.bak` backups are for.

---

## Install

Requires Python 3.9 or newer.

```bash
git clone https://github.com/kelvianlab/Excel-Mass-Replacer.git
cd Excel-Mass-Replacer
pip install -r requirements.txt
```

Or install it as a command you can run from anywhere:

```bash
pip install .
```

## Use it — desktop window

If you would rather fill in two boxes than type a command:

```bash
python -m excel_mass_replacer.gui
```

Pick the folder, type what to find and what to replace it with, press
**Preview (safe)** to see what would change, then **Replace now** to do it.

## Use it — command line

Always preview first. This changes nothing:

```bash
python -m excel_mass_replacer "PT Lama" "PT Baru" -d "C:\path\to\folder"
```

```
Scanning 128 file(s) in C:\path\to\folder
Mode: PREVIEW (dry run, nothing will be modified)

  WOULD  invoices\jan.xlsx: 42 replacement(s) in 42 cell(s) [Sheet1]
  WOULD  invoices\feb.xlsx: 17 replacement(s) in 17 cell(s) [Sheet1, Notes]

Files scanned     : 128
Files with matches: 2
Cells matched     : 59
Replacements      : 59

Preview only — nothing was written. Re-run with --apply to save changes.
Finished in 1.41s
```

Happy with it? Add `--apply`:

```bash
python -m excel_mass_replacer "PT Lama" "PT Baru" -d "C:\path\to\folder" --apply
```

Run it inside the folder itself and you can drop `-d` entirely.

### Options

| Option | What it does |
|---|---|
| `-d`, `--dir PATH` | Folder (or one file) to process. Default: the current folder. |
| `--apply` | Actually write the changes. Without it, the run is a preview. |
| `--no-backup` | Do not keep a `.bak` copy of each changed file. |
| `--no-recursive` | Stay in the given folder, do not descend into subfolders. |
| `-i`, `--ignore-case` | Match regardless of upper/lower case. |
| `-w`, `--whole-cell` | Only replace when the whole cell equals the search text. |
| `--regex` | Treat the search text as a regular expression. |
| `--include-formulas` | Also rewrite text inside `.xlsx` formulas. Ignored for `.xls`. |
| `--sheet-names` | Also rename worksheet tabs that match. |
| `--include GLOB` | Only files whose name matches, e.g. `--include "2024*.xlsx"`. Repeatable. |
| `--exclude GLOB` | Skip files whose name matches. Repeatable. |
| `--pairs-file FILE` | Run many replacements at once (see below). |
| `-j`, `--workers N` | How many files to process in parallel. Default: one per CPU core. |
| `-q`, `--quiet` | Only print the summary. |

### Many replacements in one pass

Put one `find`<kbd>Tab</kbd>`replace` rule per line in a plain text file:

```
PT Lama	PT Baru
Jl. Alamat Lama	Jl. Alamat Baru
0812-0000-0000	0812-1111-1111
```

```bash
python -m excel_mass_replacer -d ./invoices --pairs-file rules.tsv --apply
```

Lines starting with `#` are ignored. Every rule is applied in order, in a single
pass over each file.

### Use it from your own Python code

```python
from excel_mass_replacer import Rule, discover_files, run

files = discover_files("./invoices")
summary = run(files, [Rule("PT Lama", "PT Baru")], apply=True)

print(summary.total_replacements, "replacements")
for result in summary.failed_files:
    print("failed:", result.path, result.error)
```

---

## How it works

`.xlsx` and `.xlsm` files are ZIP archives full of XML; `.xls` is a binary
spreadsheet format. Both are read directly — with
[openpyxl](https://openpyxl.readthedocs.io/) and
[xlrd](https://xlrd.readthedocs.io/) / [xlwt](https://xlwt.readthedocs.io/) /
[xlutils](https://xlutils.readthedocs.io/) respectively — so no Excel process, no COM
automation, and no GUI is ever involved. Each file is written to a temporary file
first and then swapped into place atomically, so an interrupted run cannot leave you
with a half-written workbook.

## Please read this before running `--apply` on files you care about

The tool rewrites a workbook by parsing it and saving it again. Cell values, number
formats, fonts, fills, column widths and `.xlsm` macros survive that round trip, but
a few things that openpyxl cannot currently read back **are lost on any file the tool
changes**:

- charts and embedded images
- pivot tables and slicers
- cell comments / notes
- data validation in some older files

Files with **no** match are never rewritten at all, so they are never affected.

This is why backups are on by default. **Try it on a copy of your folder first**, and
keep the `.bak` files until you have checked the result. If your workbooks contain
charts or pivot tables, this tool is not the right one for them.

Two smaller notes:

- In legacy `.xls` files, text inside formulas cannot be edited; `--include-formulas`
  has no effect there.
- Replacement text is always literal, even in `--regex` mode you write the
  backreferences (`\1`) yourself — `$`, `\`, and `&` in ordinary replacement text are
  never interpreted.

## Running the tests

```bash
pip install pytest
python -m pytest tests -q
```

## License

MIT — see [LICENSE](LICENSE).
