"""Tests for the console report (intelliconverter.report) and main.py's exports.

The command-line entry point itself is tested in test_cli.py.
"""
import openpyxl
from openpyxl.utils import get_column_letter

import main
from intelliconverter import measure, report
from tests.conftest import GOLDEN_DIR, read_pdf

# Report for test.xlsx, recorded when the layout engine became the default
# (Phase 2b) and approved by the project owner.
GOLDEN_REPORT = (GOLDEN_DIR / "test_xlsx_report.txt").read_text(encoding="utf-8")


def _save(workdir, name, rows, setup=None):
    wb = openpyxl.Workbook()
    for row in rows:
        wb.active.append(row)
    if setup:
        setup(wb.active)
    wb.save(workdir / name)
    return name


def test_report_for_sample_matches_the_approved_output(workdir, capsys):
    report.analyze("test.xlsx")
    assert capsys.readouterr().out == GOLDEN_REPORT


def test_analyze_writes_output_pdf_following_the_plan(workdir, capsys):
    report.analyze("test.xlsx")
    pdf = read_pdf(workdir / "output.pdf")
    assert pdf["pages"] == 1
    assert "Page 1 of 1" in pdf["text"]


def test_explain_adds_candidates_and_costs(workdir, capsys):
    report.analyze("test.xlsx", explain=True)
    out = capsys.readouterr().out
    assert "Best candidates (lower cost is better):" in out
    assert "<- chosen" in out
    assert "Cost of the chosen layout:" in out
    assert out.startswith(GOLDEN_REPORT.split("  Part 1")[0])


def test_paper_can_be_chosen(workdir, capsys):
    report.analyze("test.xlsx", paper="A3")
    assert "  Paper   : A3 " in capsys.readouterr().out


def test_counts_and_hidden_line_for_a_sheet_with_hidden_content(workdir, capsys):
    def hide(ws):
        ws.column_dimensions["B"].hidden = True
        ws.column_dimensions["C"].hidden = True
        ws.row_dimensions[4].hidden = True

    name = _save(workdir, "hidden.xlsx",
                 [[f"r{r}c{c}" for c in range(1, 6)] for r in range(1, 5)], hide)
    report.analyze(name)
    out = capsys.readouterr().out
    assert "  Columns : 3\n" in out
    assert "  Rows    : 3\n" in out
    assert "  Hidden  : 2 columns, 1 row (not included)\n" in out


def test_no_hidden_line_when_nothing_is_hidden(workdir, capsys):
    report.analyze("test.xlsx")
    assert "Hidden" not in capsys.readouterr().out


def test_sheet_can_be_chosen(workdir, capsys):
    wb = openpyxl.Workbook()
    wb.active["A1"] = "first"
    wb.create_sheet("Second")["A1"] = "second"
    wb.save(workdir / "two.xlsx")
    report.analyze("two.xlsx", sheet="Second")
    assert "Sheet   : Second" in capsys.readouterr().out


def test_defect_1_wide_table_gets_a_layout(workdir, capsys):
    # Regression test: before Phase 2b a table wider than every page got
    # "Best layout: None". Now it is split into parts on A4.
    def widths(ws):
        for col in range(1, 31):
            ws.column_dimensions[get_column_letter(col)].width = 12

    name = _save(workdir, "wide.xlsx",
                 [[f"Heading {c}" for c in range(30)], [f"value {c}" for c in range(30)]], widths)
    report.analyze(name)
    out = capsys.readouterr().out
    assert "None" not in out
    assert "  Paper   : A4 " in out
    assert "(columns split to fit the page width)" in out


def test_defect_5_tall_sheets_get_more_pages(workdir, capsys):
    # Regression test: before Phase 2b the row count didn't affect the result.
    short = _save(workdir, "short.xlsx", [["Row", "Value"]] + [[i, "row"] for i in range(5)])
    tall = _save(workdir, "tall.xlsx", [["Row", "Value"]] + [[i, "row"] for i in range(5000)])
    report.analyze(short)
    assert "  Pages   : 1\n" in capsys.readouterr().out
    report.analyze(tall)
    pages = int(capsys.readouterr().out.split("  Pages   : ")[1].split("\n")[0].replace(",", ""))
    assert pages > 50


def test_defect_6a_pdf_contains_the_analysed_file(workdir, capsys):
    # Regression test: before Phase 1 the PDF was always built from test.xlsx.
    name = _save(workdir, "inventory.xlsx", [["Item", "Qty"], ["Widget", 7]])
    report.analyze(name)
    assert "Widget" in read_pdf(workdir / "output.pdf")["text"]


def test_main_reexports():
    assert main.analyze is report.analyze
    assert main.col_width_to_inches is measure.col_width_to_inches
    assert main.get_total_table_width is measure.get_total_table_width
