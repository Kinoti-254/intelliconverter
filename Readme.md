# IntelliConverter
An intelligent Excel-to-PDF converter that analyzes your spreadsheet 
before converting it automatically choosing the best page layout so 
tables remain readable and nothing gets awkwardly clipped.

---

## The problem it solves

Most converters blindly convert files and produce PDFs where tables are:
split across pages, text is too small to read, or content is clipped at the edges. IntelliConverter analyzes the spreadsheet dimensions first, then selects the best orientation, paper size, and scale before generating the PDF.

---

## How it works

1. Reads the Excel sheet and measures all column widths
2. Converts Excel units to inches
3. Generates layout candidates (A4 Portrait, A4 Landscape, Legal, A3...)
4. Scores each candidate based on readability and fit
5. Exports a PDF using the highest-scoring layout

---

## Installation

Make sure you have Python 3.8+ installed, then:

```bash
git clone https://github.com/Kinoti-254/intelliconverter.git
cd intelliconverter
pip install openpyxl
```

---

## Usage

```bash
python main.py
```

Place your Excel file in the project folder and update the filename 
in `main.py`. Example output:
Sheet: Sales Report

Columns: 12, Rows: 340

Total width: 11.4 inches

Recommended layout: Legal Landscape (scale: 86%)
---

## Project Structure
intelliconverter/

├── main.py          # Entry point and analysis logic

├── test.xlsx        # Sample spreadsheet for testing

└── README.md        # Documentation
---

## Roadmap

- [x] Read Excel column widths
- [x] Convert units to inches
- [ ] Generate layout candidates
- [ ] Score and select best layout
- [ ] Export PDF with chosen layout
- [ ] Conversion report (why this layout was chosen)
- [ ] FastAPI web interface

---

## Tech Stack

| Purpose            | Tool       |
|--------------------|------------|
| Read Excel         | openpyxl   |
| Generate PDF       | reportlab  |
| Layout calculation | Pure Python|
| Web API (planned)  | FastAPI    |

---

*Built as a portfolio project to explore document analysis, 
layout optimization, and file processing.*
