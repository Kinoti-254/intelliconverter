# IntelliConverter
An intelligent Excel-to-PDF converter that analyzes your spreadsheet
before converting it, automatically choosing the best page layout so
tables remain readable and nothing gets awkwardly clipped.

> **Status: early development.** The tool analyses a sheet, recommends a
> page layout and exports a basic PDF. The layout logic is simple and has
> known issues; see [Development status](#development-status).

---

## The problem it solves

Most converters blindly convert files and produce PDFs where tables are:
split across pages, text is too small to read, or content is clipped at the edges. IntelliConverter aims to analyze the spreadsheet dimensions first, then select the best orientation, paper size, and scale before generating the PDF.

---

## How it works (current version)

1. Checks the file and reads one sheet: the cells that hold values (formatting on empty cells is ignored), each column's width, and which rows and columns are hidden
2. Measures the total width of the visible columns and converts Excel units to inches
3. Compares the table width against five page layouts (A4 Portrait, A4 Landscape, A3 Portrait, Legal Landscape, A3 Landscape)
4. Scores each layout by how much of the page width the table fills, and prints a report
5. Exports `output.pdf` using the highest-scoring layout: a title, a summary line and the visible cells as a table with a repeating header row

Hidden rows and columns are left out, as Excel does when printing. Formula cells show the value Excel last saved; formulas in files that were never opened and saved in Excel have no saved value, show blank, and the report warns about them.

Only column widths are used for the layout decision. Row count, page height, font size and cell content are not used yet. In the PDF, every column gets the same width.

---

## Installation

Requires Python 3.10 or newer (tested on Python 3.10 to 3.13). Current
ReportLab releases do not run on Python 3.9.

```bash
git clone https://github.com/Kinoti-254/intelliconverter.git
cd intelliconverter
python -m venv venv
```

Activate the virtual environment:

- Windows (PowerShell): `.\venv\Scripts\Activate.ps1`
- macOS / Linux: `source venv/bin/activate`

Then install the dependencies:

```bash
pip install -r requirements.txt
```

---

## Usage

```bash
python main.py [file] [--sheet NAME] [--plan]
```

| Command | What it does |
|---------|--------------|
| `python main.py` | Analyses `test.xlsx` in the current folder |
| `python main.py sales.xlsx` | Analyses the active sheet of `sales.xlsx` |
| `python main.py sales.xlsx --sheet "Q3 Summary"` | Analyses a named sheet |
| `python main.py sales.xlsx --plan` | Previews the new layout engine's plan (see below); writes no PDF |
| `python main.py --help` | Shows the options |

The report is printed and `output.pdf` is written to the current folder,
replacing any existing `output.pdf`.

### New layout engine (preview)

`--plan` shows what the new layout engine would choose. It is a preview: the
PDF still uses the current layout until the engine becomes the default.

The engine measures every visible cell in the PDF's font and simulates the
printed table for each combination of paper (A4, Letter, Legal, A3),
orientation and font size (7 to 11 pt). Columns are sized from their content;
long text wraps; tables too wide for the page are split into column bands,
with the first column repeated in each. Each option gets a cost made of named
parts (small font, extra paper, bands, wrapping, empty space, paper type), and
the cheapest wins. The preview prints the plan, the best alternatives and the
cost breakdown, so you can see why a layout was chosen.

For sheets over 500 rows, alternatives are compared using an even sample of
rows (their page counts are marked as estimates); the chosen plan is then
worked out exactly for every row.

### Errors and limits

Problems you can fix are reported in one line, for example
`Error: File not found: sales.xlsx` (exit code 1), never as a Python traceback.
This covers missing files, unsupported or damaged files, password-protected
files, unknown sheet names, chart sheets, empty sheets, sheets where
everything is hidden, and an `output.pdf` that is open in another program.

| Limit | Value |
|-------|-------|
| File size | 20 MB |
| Rows (in the used range) | 50,000 |
| Columns (in the used range) | 100 |
| Cells (rows × columns of the used range) | 1,000,000 |

Large files are slow: on a test machine a 50,000 × 20 sheet (one million
cells) took about 20 seconds to read and about 2 minutes in total, mostly
spent drawing the 1,389-page PDF. A sheet over a limit is rejected as soon as
the limit is reached while reading, rather than after reading all of it.

Output for the included `test.xlsx`:

```
================================================
  IntelliConverter — Layout Analysis Report
================================================
  Sheet   : Sheet1
  Columns : 4
  Rows    : 3
  Width   : 4.84 inches
================================================
  Layout               Fits    Scale   Score
  ------------------------------------------
  A4 Portrait          ✅         63%      38
  A4 Landscape         ✅         45%      27
  A3 Portrait          ✅         41%      25
  Legal Landscape      ✅         37%      22
  A3 Landscape         ✅         29%      18
  ------------------------------------------

  ✨ Best layout : A4 Portrait
  📐 Scale       : 63%
  🏆 Score       : 38/100

================================================


✅ PDF saved to: output.pdf
   Layout  : A4 Portrait
   Scale   : 63%
   Rows    : 3
   Columns : 4
```

"Scale" is the share of the page width the table fills. The PDF uses it to
shrink the table's width; the font size stays the same.

---

## Running the tests

```bash
pip install -r requirements-dev.txt
pytest
```

The suite has three kinds of tests:

- **Behaviour tests** pin what the reader, the command line, the console report and the PDF (page size, page count, table contents, column widths) do, so changes to the code can be checked against them. The report for `test.xlsx` is compared byte for byte with the original program's output.
- **Regression tests** (named `test_defect_...` outside `test_known_defects.py`) keep fixed defects fixed.
- **Known-defect tests** (`tests/test_known_defects.py`) describe the correct behaviour for each open issue below. They are marked as expected failures (`xfail`) and show as `x` in the output. When a defect is fixed, its test starts passing and pytest flags it, so it can move to the regression tests.

Tests run in temporary folders, so they never change `test.xlsx` or `output.pdf` in the project folder.

---

## Project Structure

```
intelliconverter/
├── intelliconverter/          # The package
│   ├── cli.py                 # Command-line options and error messages
│   ├── errors.py              # Errors reported to the user without a traceback
│   ├── reader.py              # Checks the file and reads a sheet into a SheetModel
│   ├── measure.py             # Column widths to inches
│   ├── layout.py              # Current page layouts and the layout score
│   ├── textmetrics.py         # Measures and wraps text in the PDF font
│   ├── planner.py             # New layout engine (preview with --plan)
│   ├── report.py              # Prints the analysis report, then exports the PDF
│   └── render.py              # Builds the PDF with ReportLab
├── tests/
│   ├── golden/                # Recorded console output used by the tests
│   ├── test_cli.py
│   ├── test_reader.py
│   ├── test_measure.py
│   ├── test_layout.py
│   ├── test_textmetrics.py
│   ├── test_planner.py
│   ├── test_planner_reference.py   # Reference sheets with expected layouts
│   ├── test_report.py
│   ├── test_render.py
│   └── test_known_defects.py
├── main.py                    # Entry point (python main.py [file] [--sheet NAME] [--plan])
├── excel_to_pdf.py            # Kept for existing imports; points to render.py
├── test.xlsx                  # Sample spreadsheet
├── requirements.txt           # Runtime dependencies
├── requirements-dev.txt       # Test dependencies
├── pytest.ini                 # Test configuration
└── Readme.md
```

---

## Development status

| Capability | Status |
|------------|--------|
| Read a sheet (used range, hidden rows and columns, saved formula values) | Done |
| Input validation and size limits | Done |
| Choose the file and sheet | Done (command line) |
| Convert units to inches | Done |
| Generate layout candidates | Basic: five fixed page widths (new engine in preview: 72 candidates) |
| Score and select best layout | Basic: width fit only (new engine in preview: content, pages, bands) |
| Export PDF with chosen layout | Basic: equal column widths, fixed font |
| Conversion report (why this layout was chosen) | Not started |
| FastAPI web interface | Planned |

### Known issues

Each open issue has a matching test in `tests/test_known_defects.py`.
The numbers follow the original technical assessment.

- **1** - Tables wider than every page get no recommendation ("Best layout: None").
- **4** - Cell content is ignored, so long text in a narrow column is not accounted for.
- **5** - Row count and page height do not affect the result.

Issues 1, 4 and 5 are solved in the new layout engine (`--plan`) and will
close when it becomes the default.
- **6c** - A table too wide for every page produces a PDF labelled "Layout: None" with zero-width columns.
- **6d** - Measured column widths are ignored in the PDF; every column gets the same width.
- **6e** - Numbers are printed as Python floats, for example `23.0` instead of `23`.

Fixed (each now has a regression test): 2 hidden columns counted in the width,
3 grouped column ranges measured as one column, 6a PDF always built from
`test.xlsx`, 7 formatting on empty cells adding width, 8 formulas read as text,
9 raw errors for bad files and crashes on non-UTF-8 consoles, 6b Legal
Landscape rendered on Letter paper.

---

## Tech Stack

| Purpose            | Tool          |
|--------------------|---------------|
| Read Excel         | openpyxl      |
| Layout calculation | Pure Python   |
| Generate PDF       | reportlab     |
| Tests              | pytest, pypdf |
| Web API (planned)  | FastAPI       |

---

## License

Not yet decided.

---

*Built as a portfolio project to explore document analysis,
layout optimization, and file processing.*
