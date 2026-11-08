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

1. Reads the active sheet and measures all column widths
2. Converts Excel units to inches
3. Compares the table width against five page layouts (A4 Portrait, A4 Landscape, A3 Portrait, Legal Landscape, A3 Landscape)
4. Scores each layout by how much of the page width the table fills, and prints a report
5. Exports `output.pdf` using the highest-scoring layout: a title, a summary line and the sheet as a table with a repeating header row

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
python main.py
```

This analyses `test.xlsx` in the current folder, prints the report and writes
`output.pdf` to the current folder, replacing any existing `output.pdf`.

To analyse another file, change the filename in `main.py`. Note that the PDF
is currently always built from `test.xlsx`, whichever file is analysed
(known issue 6a below).

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

The suite has two kinds of tests:

- **Characterisation tests** pin the current behaviour of the console report and the PDF (page size, page count, table contents, column widths), so changes to the code can be checked against it.
- **Known-defect tests** (`tests/test_known_defects.py`) describe the correct behaviour for each known issue below. They are marked as expected failures (`xfail`) and show as `x` in the output. When a defect is fixed, its test starts passing and pytest flags it, so the marker can be removed.

Tests run in temporary folders, so they never change `test.xlsx` or `output.pdf` in the project folder.

---

## Project Structure

```
intelliconverter/
├── intelliconverter/          # The package
│   ├── reader.py              # Opens a workbook
│   ├── measure.py             # Column widths to inches
│   ├── layout.py              # Page layouts and the layout score
│   ├── report.py              # Prints the analysis report, then exports the PDF
│   └── render.py              # Builds the PDF with ReportLab
├── tests/
│   ├── golden/                # Recorded console output used by the tests
│   ├── test_measure.py
│   ├── test_layout.py
│   ├── test_report.py
│   ├── test_render.py
│   └── test_known_defects.py
├── main.py                    # Entry point (python main.py)
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
| Read Excel column widths | Done (active sheet only) |
| Convert units to inches | Done |
| Generate layout candidates | Basic: five fixed page widths |
| Score and select best layout | Basic: width fit only |
| Export PDF with chosen layout | Basic: equal column widths, fixed font, always from `test.xlsx` |
| Conversion report (why this layout was chosen) | Not started |
| Command-line options | Not started |
| FastAPI web interface | Planned |

### Known issues

Each issue has a matching test in `tests/test_known_defects.py`.

1. Tables wider than every page get no recommendation ("Best layout: None").
2. Hidden columns are counted in the table width.
3. Grouped column ranges (for example B:E set to one width) are measured as a single column.
4. Cell content is ignored, so long text in a narrow column is not accounted for.
5. Row count and page height do not affect the result.
6. PDF export:
    - a. The PDF is always built from `test.xlsx`, not the file being analysed.
    - b. "Legal Landscape" is rendered on Letter paper (11 in) instead of Legal (14 in).
    - c. A table too wide for every page produces a PDF labelled "Layout: None" with zero-width columns.
    - d. Measured column widths are ignored; every column gets the same width.
    - e. Numbers are printed as Python floats, for example `23.0` instead of `23`.
7. Formatting on empty cells (for example a number format far to the right) adds empty columns to the width.
8. Formula cells are read as formula text rather than their values.
9. A non-Excel file causes a raw error, and the emoji in the output crash consoles that are not UTF-8 (for example Windows output redirected to a file).

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
