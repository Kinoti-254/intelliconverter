"""Tests for the console report (intelliconverter.report) and main.py's exports.

The command-line entry point itself is tested in test_cli.py.
"""
import openpyxl
from openpyxl.utils import get_column_letter

import main
from intelliconverter import layout, measure, report
from tests.conftest import GOLDEN_DIR, read_pdf

# Console output of the original code for `python main.py` on test.xlsx,
# recorded before the Phase 0 restructure. It includes the PDF confirmation.
GOLDEN_REPORT = (GOLDEN_DIR / "test_xlsx_report.txt").read_text(encoding="utf-8")


def test_console_output_for_sample_matches_original(workdir, capsys):
    report.analyze("test.xlsx")
    assert capsys.readouterr().out == GOLDEN_REPORT


def test_analyze_writes_output_pdf_in_current_folder(workdir, capsys):
    report.analyze("test.xlsx")
    assert (workdir / "output.pdf").is_file()


def test_recommends_smallest_page_the_table_fits(workdir, save_workbook, capsys):
    # 10 columns of width 11 = 8.54 in: too wide for A4 Portrait (7.7 in)
    wb = openpyxl.Workbook()
    ws = wb.active
    for col in range(1, 11):
        ws.cell(row=1, column=col, value="x")
        ws.column_dimensions[get_column_letter(col)].width = 11
    report.analyze(save_workbook(wb))
    assert "Best layout : A4 Landscape" in capsys.readouterr().out


def test_counts_and_hidden_line_for_a_sheet_with_hidden_content(workdir, save_workbook, capsys):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in range(1, 5):
        ws.append([f"r{row}c{col}" for col in range(1, 6)])
    ws.column_dimensions["B"].hidden = True
    ws.column_dimensions["C"].hidden = True
    ws.row_dimensions[4].hidden = True
    report.analyze(save_workbook(wb))
    out = capsys.readouterr().out
    assert "  Columns : 3\n" in out
    assert "  Rows    : 3\n" in out
    assert "  Hidden  : 2 columns, 1 row (not included)\n" in out


def test_no_hidden_line_when_nothing_is_hidden(workdir, capsys):
    report.analyze("test.xlsx")
    assert "Hidden" not in capsys.readouterr().out


def test_sheet_can_be_chosen(workdir, save_workbook, capsys):
    wb = openpyxl.Workbook()
    wb.active["A1"] = "first"
    wb.create_sheet("Second")["A1"] = "second"
    report.analyze(save_workbook(wb), sheet="Second")
    assert "Sheet   : Second" in capsys.readouterr().out


def test_defect_6a_pdf_contains_the_analysed_file(workdir, capsys):
    # Regression test: before Phase 1 the PDF was always built from test.xlsx.
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Inventory"
    ws.append(["Item", "Qty"])
    ws.append(["Widget", 7])
    wb.save(workdir / "inventory.xlsx")
    report.analyze("inventory.xlsx")
    text = read_pdf(workdir / "output.pdf")["text"]
    assert "Inventory" in text and "Widget" in text


def test_main_reexports_the_original_names():
    assert main.analyze is report.analyze
    assert main.score_layout is layout.score_layout
    assert main.LAYOUTS is layout.LAYOUTS
    assert main.col_width_to_inches is measure.col_width_to_inches
    assert main.get_total_table_width is measure.get_total_table_width
