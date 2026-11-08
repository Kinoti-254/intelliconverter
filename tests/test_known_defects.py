"""Known defects, captured as expected failures.

Each test asserts the CORRECT behaviour, so it fails today and is marked
xfail. pytest.ini sets xfail_strict = true: when a later phase fixes a
defect, its test passes, pytest reports that as a failure, and the test
moves to the relevant test file as a regression test.

Numbering follows the technical assessment (defect 6 became 6a-6e, the
PDF defects). Fixed in Phase 1, now regression tests:
    2, 3, 7, 8 -> test_reader.py
    6a         -> test_report.py
    9a, 9b     -> test_cli.py
Still open: 1, 4, 5 (layout engine, Phase 2) and 6b-6e (PDF, Phase 3).
"""
import openpyxl
import pytest
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import LEGAL, landscape

from intelliconverter import render
from intelliconverter.measure import col_width_to_inches, get_total_table_width
from intelliconverter.reader import load_sheet
from intelliconverter.report import analyze
from tests.conftest import SAMPLE_XLSX, read_pdf


def _sheet_with_widths(widths):
    wb = openpyxl.Workbook()
    ws = wb.active
    for col, width in enumerate(widths, start=1):
        ws.cell(row=1, column=col, value="x")
        ws.column_dimensions[get_column_letter(col)].width = width
    return wb


# --- Layout engine (Phase 2) ------------------------------------------------

@pytest.mark.xfail(reason="Defect 1: tables wider than every page get no recommendation")
def test_defect_1_wide_table_still_gets_a_layout(workdir, save_workbook, capsys):
    # 30 columns of width 12 = 27.8 in, wider than A3 Landscape (16.5 in).
    # Expected: some layout is recommended (smaller font or column bands).
    analyze(save_workbook(_sheet_with_widths([12] * 30)))
    assert "Best layout : None" not in capsys.readouterr().out


@pytest.mark.xfail(reason="Defect 4: cell content is ignored when measuring columns")
def test_defect_4_long_text_widens_its_column(save_workbook):
    wb = openpyxl.Workbook()
    wb.active["A1"] = "Nairobi Regional Sales Office"  # 29 characters
    model = load_sheet(save_workbook(wb))
    # A default-width column holds about 8 characters; the text needs more room.
    assert get_total_table_width(model) > col_width_to_inches(8.43)


@pytest.mark.xfail(reason="Defect 5: row count and page height do not affect the analysis")
def test_defect_5_tall_sheet_is_analysed_differently(workdir, save_workbook, capsys):
    def two_column_sheet(rows):
        wb = openpyxl.Workbook()
        ws = wb.active
        for row in range(1, rows + 1):
            ws.cell(row=row, column=1, value=row)
            ws.cell(row=row, column=2, value="row")
        return wb

    def without_row_counts(text):
        return [line for line in text.splitlines() if "Rows" not in line]

    analyze(save_workbook(two_column_sheet(5), "short.xlsx"))
    short = capsys.readouterr().out
    analyze(save_workbook(two_column_sheet(5000), "tall.xlsx"))
    tall = capsys.readouterr().out
    # Expected: page count or pagination differs, not just the row counts.
    assert without_row_counts(short) != without_row_counts(tall)


# --- PDF export (Phase 3) ---------------------------------------------------

@pytest.mark.xfail(reason="Defect 6b: Legal Landscape is rendered on Letter paper")
def test_defect_6b_legal_landscape_uses_legal_paper():
    # Legal is 8.5 x 14 in; Letter is 8.5 x 11 in.
    assert render.get_page_size("Legal Landscape") == landscape(LEGAL)


@pytest.mark.xfail(reason="Defect 6c: a table too wide for every page produces zero-width columns")
def test_defect_6c_wide_table_pdf_is_usable(workdir, captured_tables, capsys):
    _sheet_with_widths([12] * 30).save(workdir / "wide.xlsx")
    analyze("wide.xlsx")
    text = read_pdf(workdir / "output.pdf")["text"]
    assert "Layout: None" not in text
    assert all(width > 0 for width in captured_tables[0]["col_widths"])


@pytest.mark.xfail(reason="Defect 6d: measured column widths are ignored; every column is equal")
def test_defect_6d_column_widths_follow_the_sheet(tmp_path, captured_tables, capsys):
    # In test.xlsx column B (25.63) is about twice as wide as A, C and D (12.63).
    render.build_pdf(SAMPLE_XLSX, str(tmp_path / "out.pdf"), "A4 Portrait", 63)
    widths = captured_tables[0]["col_widths"]
    assert widths[1] > widths[0]


@pytest.mark.xfail(reason="Defect 6e: numbers are printed as Python floats, e.g. 23.0")
def test_defect_6e_whole_numbers_print_without_decimals(tmp_path, capsys):
    out = tmp_path / "out.pdf"
    render.build_pdf(SAMPLE_XLSX, str(out), "A4 Portrait", 63)
    text = read_pdf(out)["text"]
    # Excel shows the age as 23 (General format).
    assert "23" in text and "23.0" not in text
