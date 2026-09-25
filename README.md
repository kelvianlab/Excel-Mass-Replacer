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

### Step by step

1. **Click `Browse…` and pick the folder that holds your Excel files.**
   The box does not start on the right folder — it opens wherever you launched
   the app from, which is usually the app's own folder. That folder contains no
   spreadsheets, so a run started there just reports
   *"No .xlsx, .xlsm or .xls files found in…"*. Point it at your own files.
2. **Type the text in `Find`, and what should take its place in `Replace with`.**
   Leave `Replace with` empty to delete the text instead of replacing it.
3. **Press `Preview (safe)` first.** Nothing is written. The log lists every file
   that would change and how many replacements each one would get.
4. **Read the numbers before you commit to them.** If the count is far higher
   than you expected, your search text is matching more than you meant — see
   `Whole cell only` below, adjust, and preview again.
5. **Press `Replace now`.** You will get a confirmation dialog showing the folder,
   the two texts, and whether backups are on. Nothing happens until you accept it.

### The four checkboxes, with a worked example

The defaults are the safe ones. You only need to change them for specific jobs.

Say one spreadsheet holds these five cells, and you search for `PT Lama` and
replace it with `PT Baru`:

| | A |
|---|---|
| **1** | `PT Lama Jaya` |
| **2** | `pt lama` |
| **3** | `Invoice for PT Lama` |
| **4** | `PT LAMA` |
| **5** | `PT Lama` |

Here is what each setting actually does to those cells. These are real results
from running the tool, not an illustration:

| Cell before | Default | `Ignore case` | `Whole cell only` | Both on |
|---|---|---|---|---|
| `PT Lama Jaya` | **PT Baru** Jaya | **PT Baru** Jaya | — | — |
| `pt lama` | — | **PT Baru** | — | **PT Baru** |
| `Invoice for PT Lama` | Invoice for **PT Baru** | Invoice for **PT Baru** | — | — |
| `PT LAMA` | — | **PT Baru** | — | **PT Baru** |
| `PT Lama` | **PT Baru** | **PT Baru** | **PT Baru** | **PT Baru** |
| **Replacements** | **3** | **5** | **1** | **3** |

A dash means the cell was left exactly as it was.

#### `Ignore case` — whether capitals matter

- **Off (default):** only `PT Lama` matches. `pt lama` and `PT LAMA` are skipped,
  because the capitals differ.
- **On:** all three spellings count as the same text, so all of them are replaced.
  That is why the total goes from 3 to 5.

Turn it on when the data was typed by different people and the capitalisation is
inconsistent.

#### `Whole cell only` — whether the cell must match exactly

- **Off (default):** the text is found *inside* the cell. `PT Lama Jaya` is
  replaced, because it contains `PT Lama`.
- **On:** a cell is replaced only when its entire contents are exactly
  `PT Lama`, with nothing before or after. Only cell A5 qualifies, so the total
  drops to 1.

Turn it on when your search text is also part of longer values you must not
touch. Searching for `Jakarta` with this off rewrites `Jakarta Selatan` too;
with it on, that cell is left alone.

#### `Include subfolders` — how deep the search goes

Given this layout, with the folder picker pointed at `invoices`:

```
invoices/
├── jan.xlsx
├── feb.xlsx
└── archive/
    └── 2023.xlsx
```

- **On (default):** all three files are processed, `archive/2023.xlsx` included.
- **Off:** only `jan.xlsx` and `feb.xlsx`; the `archive` folder is skipped.

#### `Keep .bak backup` — your way back

- **On (default):** every file that changes gets a copy saved beside it first, so
  the folder ends up with `jan.xlsx` (changed) and `jan.xlsx.bak` (the original).
  To undo, delete the changed file and rename the `.bak` back.
- **Off:** no copies are kept and the change cannot be undone from inside the tool.

Files with no match are never rewritten, so they never get a `.bak` either.

### On your first run

Copy a handful of your spreadsheets into an empty folder and run the tool there
first. Once you have seen it do the right thing on those, point it at the real
folder. This is worth the two minutes for any tool that edits many files at once.

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

`-d` must point at the folder holding your spreadsheets, not at the folder you
unpacked this tool into. Leave the replacement empty (`""`) to delete the search
text rather than replace it. The same first-run advice applies here: try it on a
copy of a few files before pointing it at the real folder.

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
