# IntelliConverter
An intelligent Excel-to-PDF converter that analyzes your spreadsheet
before converting it, automatically choosing the best page layout so
tables remain readable and nothing gets awkwardly clipped.

> **Status: in development.** The layout engine and PDF export work from the
> command line, for one sheet or a whole workbook. A visual preview and a web
> API are planned; see [Development status](#development-status).

---

## The problem it solves

Most converters blindly convert files and produce PDFs where tables are:
split across pages, text is too small to read, or content is clipped at the edges. IntelliConverter analyzes the spreadsheet's content first, then selects the paper, orientation, font size and column widths that make the most readable document.

---

## How it works

1. **Reads the sheets to convert** (by default every visible worksheet): only cells that hold values (formatting on empty cells is ignored), which rows and columns are hidden, and each cell's number format. Hidden sheets, chart sheets and empty sheets are skipped and listed in the report.
2. **Shows values as Excel does:** thousands separators, decimals, percentages, currency symbols and dates (for example `1,234.50`, `25%`, `08/11/2026`).
3. **Measures every visible cell** in the PDF's font (Helvetica).
4. **Simulates the printed table** for every combination of paper (A4, Letter, Legal, A3), orientation (portrait and landscape) and font size (7 to 11 pt):
    - columns are sized from their content and headers; long text wraps;
    - a table too wide for the page is split into parts (column bands), each repeating the first column;
    - rows are paginated without splitting a row across pages, with the header repeated on every page.
5. **Chooses the cheapest option** using a cost made of named parts: small font, extra paper, column parts, wrapping, empty space and paper type. Neither orientation is preferred in advance: portrait or landscape wins only through these costs. A4 is the standard paper: Legal and A3 cost more than an extra column part, so they are chosen only when nothing else fits, or when you ask for them.
6. **Draws the PDF exactly as planned:** same widths, line breaks, row heights and page breaks, a title on the first page and page numbers on every page.

**Each sheet is planned independently**, so one PDF can mix paper sizes and
orientations: a wide sheet on landscape pages, followed by a narrow one on
portrait pages. Each sheet starts on a new page with its own title and a PDF
bookmark, and the footer shows both the sheet's page number and the
document's.

Hidden rows and columns are left out, as Excel does when printing. Formula
cells show the value Excel last saved; formulas in files that were never
opened and saved in Excel show blank, and the report warns about them.

For sheets over 500 rows, the options are compared using an even sample of
rows; the chosen layout is then worked out exactly for every row.

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
python main.py [file] [--sheet NAME ...] [--paper PAPER] [--explain] [-o PDF]
```

| Command | What it does |
|---------|--------------|
| `python main.py` | Converts every visible sheet of `test.xlsx` in the current folder |
| `python main.py sales.xlsx` | Converts every visible sheet of `sales.xlsx` into one PDF |
| `python main.py sales.xlsx --sheet "Q3 Summary"` | Converts one named sheet |
| `python main.py sales.xlsx --sheet Costs --sheet Sales` | Converts named sheets, in that order |
| `python main.py sales.xlsx -o sales.pdf` | Writes `sales.pdf` instead of `output.pdf` |
| `python main.py sales.xlsx --paper A3` | Uses A3 paper for every sheet (also `A4`, `Letter`, `Legal`; default `auto`). Both orientations are still compared. |
| `python main.py sales.xlsx --explain` | Also shows each sheet's best alternatives and the cost of its layout |
| `python main.py --help` | Shows the options |

The report is printed and the PDF (`output.pdf` unless `-o` says otherwise)
is written, replacing any existing file of that name.

Output for the included `test.xlsx`:

```
================================================
  IntelliConverter — Layout Report
================================================
  Sheet   : Sheet1
  Columns : 4
  Rows    : 3
================================================
  Paper   : A4 Portrait
  Font    : 10 pt
  Parts   : 1
  Pages   : 1
  Part 1  : A 0.54, B 0.44, C 0.78, D 0.91 (inches)
================================================

✅ PDF saved to: output.pdf
```

With several sheets, the report shows one block per sheet and a total, for
example:

```
  Sheets  : 2 (Quarterly Figures, Stock List)
  Skipped : Scratch (empty)
================================================
  Sheet   : Quarterly Figures
  ...
  Paper   : A4 Landscape
  Pages   : 2
  ----------------------------------------------
  Sheet   : Stock List
  ...
  Paper   : A4 Portrait
  Pages   : 4
================================================
  Total   : 2 sheets, 6 pages
```

With `--explain`, the report also lists the best alternatives and the cost
breakdown, for example:

```
  Best candidates (lower cost is better):
  Candidate           Font  Parts   Pages    Cost
  A4 Portrait           10      1       1    1.26  <- chosen
  A4 Landscape          10      1       1    1.50
  ...
  Cost of the chosen layout:
    empty_width        1.26
    total              1.26
```

### Errors and limits

Problems you can fix are reported in one line, for example
`Error: File not found: sales.xlsx` (exit code 1), never as a Python traceback.
This covers missing files, unsupported or damaged files, password-protected
files, unknown sheet names, a named sheet that is a chart sheet, empty or
entirely hidden, a workbook with nothing to convert, a cell too long to fit
on any page (the message names the sheet), and a PDF that is open in another
program. When any sheet fails, no PDF is written.

| Limit | Value |
|-------|-------|
| File size | 20 MB |
| Rows (in the used range) | 50,000 |
| Columns (in the used range) | 100 |
| Cells (rows × columns of the used range, all selected sheets together) | 1,000,000 |

A sheet over a limit is rejected as soon as the limit is reached while
reading. On a test machine a 50,000 × 20 sheet (one million cells) took about
70 seconds in total (reading 21 s, planning 18 s, drawing the 4,168-page PDF
30 s) and about 1 GB of memory.

---

## Running the tests

```bash
pip install -r requirements-dev.txt
pytest
```

The suite covers:

- **Behaviour tests** for the reader, number formats, text measurement, the layout engine, the command line and the report. The report for `test.xlsx` is compared byte for byte with an approved copy.
- **PDF tests** that read each generated PDF back and check every promise the plan makes: page count and size, column widths, row heights, wrapped lines, page breaks, the repeated header, page numbers, and that every visible cell appears once. Multi-sheet PDFs are checked page by page against each sheet's own plan, including bookmarks.
- **Reference sheets** (`tests/test_planner_reference.py`) with the layout a person would expect, written before the engine's costs were tuned.
- **Regression tests** (named `test_defect_...`) for every defect found in the original technical assessment.

Known defects can be recorded as tests marked `xfail` (expected to fail);
`pytest.ini` makes them strict, so a fixed defect is flagged until its test
becomes a regression test. There are none open at present.

Tests run in temporary folders, so they never change `test.xlsx` or `output.pdf` in the project folder.

---

## Project Structure

```
intelliconverter/
├── intelliconverter/          # The package
│   ├── cli.py                 # Command-line options and error messages
│   ├── errors.py              # Errors reported to the user without a traceback
│   ├── reader.py              # Checks the file and reads a sheet into a SheetModel
│   ├── numberformat.py        # Shows values the way Excel does
│   ├── textmetrics.py         # Measures and wraps text in the PDF font
│   ├── planner.py             # The layout engine
│   ├── report.py              # Prints the report, then exports the PDF
│   ├── render.py              # Draws the PDF exactly as planned
│   └── measure.py             # Excel column widths to inches (for a future mode)
├── tests/
│   ├── golden/                # Approved report output for test.xlsx
│   ├── test_cli.py
│   ├── test_reader.py
│   ├── test_numberformat.py
│   ├── test_textmetrics.py
│   ├── test_planner.py
│   ├── test_planner_reference.py   # Reference sheets with expected layouts
│   ├── test_measure.py
│   ├── test_report.py
│   └── test_render.py
├── main.py                    # Entry point
├── excel_to_pdf.py            # Kept for code written before the layout engine
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
| Number and date formats (basic) | Done |
| Content-based column widths, wrapping and font sizing | Done |
| Paper and orientation selection (A4, Letter, Legal, A3) | Done |
| Pagination with repeated headers; no row split across pages | Done |
| Wide tables split into parts that repeat the first column | Done |
| PDF that follows the plan exactly | Done |
| Explanation of the chosen layout (`--explain`) | Done |
| Several sheets in one PDF, each with its own layout | Done |
| Preserve Excel's column widths (optional mode) | Planned |
| Visual preview and user adjustments | Planned |
| FastAPI web interface | Planned |

### Limitations

- The first visible row of each sheet is always treated as its header.
- `--paper` applies to every sheet; there is no per-sheet paper choice yet.
- The first column is assumed to identify rows when a table is split into parts.
- Fonts, colours, borders, merged cells, images and charts from the sheet are not reproduced; the PDF uses its own clean table style.
- Number formats cover the common cases; fractions and some special formats fall back to Excel's General format.
- Large sheets are slow and memory-heavy (see the limits above).

All defects from the original technical assessment are fixed; each has a
regression test.

---

## Tech Stack

| Purpose            | Tool          |
|--------------------|---------------|
| Read Excel         | openpyxl      |
| Layout engine      | Pure Python   |
| Generate PDF       | reportlab     |
| Tests              | pytest, pypdf |
| Web API (planned)  | FastAPI       |

---

## License

Not yet decided.

---

*Built as a portfolio project to explore document analysis,
layout optimization, and file processing.*
