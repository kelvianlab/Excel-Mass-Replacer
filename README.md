# Excel Mass Replacer

A lightning-fast, headless Python utility to mass find-and-replace text across all
Excel files (`.xlsx`, `.xlsm`, `.xls`) in a directory — without opening Microsoft Excel.

Excel VBA macros drive a real Excel process through COM. That is why they pop
windows open, hang on a dialog nobody can see, eat memory, and die halfway through a
folder. This tool never touches Excel at all: it reads and rewrites the workbooks as
pure data, so the run is silent, repeatable, and finishes in seconds.

**Measured on this project's own benchmark:** 50 workbooks, 300,000 matching cells,
replaced and saved in **5.9 seconds** on a 4-core machine.

**On Windows you do not need Python at all** — download
[`ExcelMassReplacer-App.exe`](https://github.com/kelvianlab/Excel-Mass-Replacer/releases/latest)
and double-click it.

---

## What it does

- Replaces a word or sentence with another one in **every** Excel file in a folder
  (and its subfolders, unless you say otherwise).
- Handles `.xlsx` and `.xlsm` (macro-enabled, macros kept) with nothing else in
  the file disturbed. Legacy `.xls` works too, but is rebuilt rather than edited
  in place — see [the warning below](#please-read-this-before-running---apply-on-files-you-care-about)
  before using it on anything you cannot replace.
- **Previews by default.** Nothing is written until you explicitly ask for it.
- Copies every file it changes into one backup folder first, under its own name,
  unless you turn that off.
- **Leaves the rest of the workbook exactly as it found it** in `.xlsx`/`.xlsm`:
  the results Excel cached for each formula, images, charts, comments, form
  controls and printer settings all survive untouched.
- Rewrites the choices in a dropdown list when they are typed into the rule
  itself, and tells you which files that happened in.
- Skips Excel's `~$` lock files, and keeps going when one workbook is corrupt
  instead of aborting the whole batch.
- Runs every CPU core in parallel.
- Never launches Excel, never shows a dialog, never waits for a click.
- Ships as a single Windows `.exe` that needs no Python and no installation.

## What it does not do

- It does not edit Word, PowerPoint, CSV, or ODS files.
- It does not change numbers, dates, or formula *results* — only text. The value
  Excel last calculated for a formula is carried over untouched, so a program
  reading the file afterwards still sees it.
- It does not recalculate anything. It has no formula engine and never needs one,
  because it does not disturb the results that are already in the file.
- It does not undo a run for you. That is what the backup folder is for.

---

## Download for Windows — no Python, no typing

Go to the
[latest release](https://github.com/kelvianlab/Excel-Mass-Replacer/releases/latest).
Under **Assets** you will see four files. This is what each one is:

| File | What it is | Do you want it? |
|---|---|---|
| **`ExcelMassReplacer-App.exe`** | The app. Double-click it and a window opens. | ✅ **Yes — download this one** |
| `ExcelMassReplacer-CommandLine.exe` | Same tool, typed commands instead of a window. | Only if you script or schedule things |
| `Source code (zip)` | The programming files. Added automatically by GitHub. | No |
| `Source code (tar.gz)` | The same files, in the format Mac and Linux prefer. | No |

**In short: download `ExcelMassReplacer-App.exe` and ignore the rest.**

Nothing to install and nothing to set up — Python and every library it needs are
already inside that one file. Put it anywhere you like (Desktop, a USB stick, a
shared drive) and it runs from there.

### Windows will warn you the first time

The executables are not code-signed — a signing certificate costs a few hundred
dollars a year, which this project does not have. So Windows SmartScreen shows
**"Windows protected your PC"** the first time you run it. That warning means
"this file is not signed and not yet widely downloaded", not "this file is
harmful".

To continue: click **More info**, then **Run anyway**. You will only be asked
once per machine.

If you would rather not take that on trust, you do not have to. Every release is
built by GitHub Actions from the source in this repository — you can read the
[build workflow](.github/workflows/release.yml), see the exact commit it was
built from on the release page, and build it yourself with the steps below.

### What is in the file

| | |
|---|---|
| Size | about 13 MB — that is Python plus the Excel libraries bundled in |
| Needs | Windows 10 or 11, 64-bit. Nothing else |
| Internet | never used. The tool only touches the folder you point it at |
| Microsoft Excel | not required, and never launched |

## Install from source

For macOS and Linux, or if you would rather run the Python directly.
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

### Build the Windows executable yourself

```bash
pip install pyinstaller
pyinstaller --clean --noconfirm --onefile --windowed --name ExcelMassReplacer-App --paths . packaging/gui_entry.py
```

The result lands in `dist/`. Run this on Windows: PyInstaller builds for the
system it runs on, so a Windows `.exe` has to be built on Windows.

## Use it — desktop window

Double-click **`ExcelMassReplacer-App.exe`**, or from a source checkout:

```bash
python -m excel_mass_replacer.gui
```

Either way you get the same window: fill in two boxes and press a button.

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

Say one spreadsheet holds these five cells, and you search for `Acme` and
replace it with `Globex`:

| | A |
|---|---|
| **1** | `Acme Holdings` |
| **2** | `acme` |
| **3** | `Invoice for Acme` |
| **4** | `ACME` |
| **5** | `Acme` |

Here is what each setting actually does to those cells. These are real results
from running the tool, not an illustration:

| Cell before | Default | `Ignore case` | `Whole cell only` | Both on |
|---|---|---|---|---|
| `Acme Holdings` | **Globex** Holdings | **Globex** Holdings | — | — |
| `acme` | — | **Globex** | — | **Globex** |
| `Invoice for Acme` | Invoice for **Globex** | Invoice for **Globex** | — | — |
| `ACME` | — | **Globex** | — | **Globex** |
| `Acme` | **Globex** | **Globex** | **Globex** | **Globex** |
| **Replacements** | **3** | **5** | **1** | **3** |

A dash means the cell was left exactly as it was.

#### `Ignore case` — whether capitals matter

- **Off (default):** only `Acme` matches. `acme` and `ACME` are skipped,
  because the capitals differ.
- **On:** all three spellings count as the same text, so all of them are replaced.
  That is why the total goes from 3 to 5.

Turn it on when the data was typed by different people and the capitalisation is
inconsistent.

#### `Whole cell only` — whether the cell must match exactly

- **Off (default):** the text is found *inside* the cell. `Acme Holdings` is
  replaced, because it contains `Acme`.
- **On:** a cell is replaced only when its entire contents are exactly
  `Acme`, with nothing before or after. Only cell A5 qualifies, so the total
  drops to 1.

Turn it on when your search text is also part of longer values you must not
touch. Searching for `York` with this off turns `New York` into `New Boston`;
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

#### `Keep a backup copy` — your way back

- **On (default):** before anything is written, every file that is about to change
  is copied into one folder named `Excel-Mass-Replacer-Backup`, inside a sub-folder
  named for the date and time of the run. Each copy keeps its real name and its
  place in your sub-folders, so the backup is a plain mirror of the originals:

  ```
  PO 2026/
    jan.xlsx                     <- changed
    archive/feb.xlsx             <- changed
    Excel-Mass-Replacer-Backup/
      20261004-210145/
        jan.xlsx                 <- the original, openable as it is
        archive/feb.xlsx
  ```

  To undo, copy the files back over the changed ones. There is nothing to rename.
- **Off:** no copies are kept and the change cannot be undone from inside the tool.

Files with no match are never rewritten, so they are never copied either. A later
run skips the backup folder, so your backups are never themselves replaced.

### On your first run

Copy a handful of your spreadsheets into an empty folder and run the tool there
first. Once you have seen it do the right thing on those, point it at the real
folder. This is worth the two minutes for any tool that edits many files at once.

## Use it — command line

Always preview first. This changes nothing:

```bash
python -m excel_mass_replacer "Acme" "Globex" -d "C:\path\to\folder"
```

Downloaded the Windows build instead? Use `ExcelMassReplacer-CommandLine.exe` in
place of `python -m excel_mass_replacer` — every option below is identical.

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
python -m excel_mass_replacer "Acme" "Globex" -d "C:\path\to\folder" --apply
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
| `--no-backup` | Do not copy the originals anywhere before changing them. |
| `--backup-dir PATH` | Put the copies here instead of in a timestamped `Excel-Mass-Replacer-Backup` folder inside the scanned folder. |
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
Acme	Globex
12 Old Street	480 New Avenue
+1-555-0100	+1-555-0199
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
summary = run(files, [Rule("Acme", "Globex")], apply=True)

print(summary.total_replacements, "replacements")
for result in summary.failed_files:
    print("failed:", result.path, result.error)
```

---

## How it works

`.xlsx` and `.xlsm` files are ZIP archives full of XML. The tool opens the archive
itself, using nothing but the Python standard library, and edits the XML parts that
hold text: the shared string table, any inline strings, the typed-in choices of a
dropdown list, and formulas when you ask for those. A part with nothing to replace
is copied across byte for byte, which is why everything else in the workbook comes
out identical. That is a deliberate choice over loading the workbook into a library
and saving it again, since saving it again means rebuilding it from whatever the
library understands and quietly losing the rest.

`.xls` is a binary format from an older Excel and is handled separately, through
[xlrd](https://xlrd.readthedocs.io/) / [xlwt](https://xlwt.readthedocs.io/) /
[xlutils](https://xlutils.readthedocs.io/). That path does rebuild the workbook.

No Excel process, no COM automation, and no GUI is ever involved. Each file is
written to a temporary file first and then swapped into place atomically, so an
interrupted run cannot leave you with a half-written workbook.

## Please read this before running `--apply` on files you care about

An `.xlsx` or `.xlsm` file is a ZIP of XML parts. This tool opens that ZIP, rewrites
only the text nodes that actually change, and copies every other byte straight
through. Anything it does not understand is therefore carried over exactly: the
value Excel cached for each formula, images, charts, pivot tables, comments, form
controls, printer settings and macros.

It did not always work this way. Until this was fixed, the tool loaded each workbook
through openpyxl and saved it again, which rebuilt the file from what openpyxl could
model and silently dropped the rest. The text came out right, so the damage was easy
to miss: a purchase order would keep its wording and lose its logo, and every formula
kept its formula while losing the number Excel had worked out, which is the number
every *other* program reads. Files looked fine opened by hand in Excel, because Excel
recalculates on open, and came back empty to anything reading them automatically.

**Legacy `.xls` is still the old story.** Those files go through xlrd/xlwt, which
does rebuild the workbook, so the caveats above still apply to `.xls` only.

Files with **no** match are never rewritten at all, so they are never affected.

Backups are still on by default, and **trying it on a copy of your folder first is
still the right habit.**

### Dropdown lists

A dropdown whose choices come from cells (`$A$1:$A$9`, or `Supplier!$A$2:$A$371` on
another sheet) needs nothing special: its text lives in those cells and is replaced
along with everything else, so the dropdown follows automatically.

A dropdown whose choices are **typed into the validation rule itself** is different.
That text is part of a rule, not a cell, so the tool rewrites it and then says so in
the report, naming each file. Treat that line as something to check: changing it
changes what people are allowed to enter, not just what they can read. Leave it
unchanged and you get the opposite problem, a cell holding a value its own dropdown
no longer offers, which only shows up when somebody next uses that dropdown.

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
