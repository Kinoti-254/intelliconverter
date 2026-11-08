"""Known defects, captured as expected failures.

Each test asserts the CORRECT behaviour, so it fails today and is marked
xfail. pytest.ini sets xfail_strict = true: when a later phase fixes a
defect, its test passes, pytest reports that as a failure, and the marker
is removed. No test here asserts the wrong behaviour as if it were right.

Numbering follows the technical assessment. Defect 6 ("no PDF is
generated") did not apply to this version of the code, which does export
a PDF; it is replaced by the PDF defects 6a-6e found in that export.

Some tests call today's functions. When later phases replace those
functions, the tests move to the new API with the same expected behaviour.
"""
import os
import shutil
import subprocess
import sys
import zipfile

import openpyxl
import pytest
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import LEGAL, landscape

from intelliconverter import render
from intelliconverter.measure import col_width_to_inches, get_total_table_width
from intelliconverter.reader import load_active_sheet
from intelliconverter.report import analyze
from tests.conftest import REPO_ROOT, SAMPLE_XLSX, read_pdf


def _sheet_with_widths(widths):
    wb = openpyxl.Workbook()
    ws = wb.active
    for col, width in enumerate(widths, start=1):
        ws.cell(row=1, column=col, value="x")
        ws.column_dimensions[get_column_letter(col)].width = width
    return wb


# --- Analysis defects -------------------------------------------------------

@pytest.mark.xfail(reason="Defect 1: tables wider than every page get no recommendation")
def test_defect_1_wide_table_still_gets_a_layout(workdir, save_workbook, capsys):
    # 30 columns of width 12 = 27.8 in, wider than A3 Landscape (16.5 in).
    # Expected: some layout is recommended (smaller font or column bands).
    analyze(save_workbook(_sheet_with_widths([12] * 30)))
    assert "Best layout : None" not in capsys.readouterr().out


@pytest.mark.xfail(reason="Defect 2: hidden columns are counted in the table width")
def test_defect_2_hidden_columns_are_excluded(reload_sheet):
    wb = _sheet_with_widths([20] * 10)
    for letter in "CDEFGH":
        wb.active.column_dimensions[letter].hidden = True
    ws = reload_sheet(wb)
    # Only A, B, I and J are visible.
    assert get_total_table_width(ws) == pytest.approx(4 * col_width_to_inches(20))


@pytest.mark.xfail(reason="Defect 3: a grouped column range is counted as one column")
def test_defect_3_grouped_column_range_counts_every_column(reload_sheet):
    wb = openpyxl.Workbook()
    ws = wb.active
    for col in range(1, 6):
        ws.cell(row=1, column=col, value="x")
    ws.column_dimensions.group("B", "E", hidden=False)
    ws.column_dimensions["B"].width = 30
    ws = reload_sheet(wb)
    # openpyxl stores B:E as one entry keyed "B"; C, D and E are also width 30.
    expected = col_width_to_inches(8.43) + 4 * col_width_to_inches(30)
    assert get_total_table_width(ws) == pytest.approx(expected)


@pytest.mark.xfail(reason="Defect 4: cell content is ignored when measuring columns")
def test_defect_4_long_text_widens_its_column(reload_sheet):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Nairobi Regional Sales Office"  # 29 characters
    ws = reload_sheet(wb)
    # A default-width column holds about 8 characters; the text needs more room.
    assert get_total_table_width(ws) > col_width_to_inches(8.43)


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


# --- PDF export defects (replace the assessment's defect 6) -----------------

@pytest.mark.xfail(reason="Defect 6a: the PDF is always built from test.xlsx, not the analysed file")
def test_defect_6a_pdf_contains_the_analysed_file(workdir, capsys):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Inventory"
    ws.append(["Item", "Qty"])
    ws.append(["Widget", 7])
    wb.save(workdir / "inventory.xlsx")
    analyze("inventory.xlsx")
    text = read_pdf(workdir / "output.pdf")["text"]
    assert "Inventory" in text and "Widget" in text


@pytest.mark.xfail(reason="Defect 6b: Legal Landscape is rendered on Letter paper")
def test_defect_6b_legal_landscape_uses_legal_paper():
    # Legal is 8.5 x 14 in; Letter is 8.5 x 11 in.
    assert render.get_page_size("Legal Landscape") == landscape(LEGAL)


@pytest.mark.xfail(reason="Defect 6c: a table too wide for every page produces zero-width columns")
def test_defect_6c_wide_table_pdf_is_usable(workdir, captured_tables, capsys):
    _sheet_with_widths([12] * 30).save(workdir / "test.xlsx")
    analyze("test.xlsx")
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


# --- Input and output defects -----------------------------------------------

@pytest.mark.xfail(reason="Defect 7: formatting on empty cells inflates the measured width")
def test_defect_7_formatted_empty_cells_do_not_add_width(reload_sheet):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "data"
    ws["Z1"].number_format = "0.00"  # formatted, but empty
    ws = reload_sheet(wb)
    assert get_total_table_width(ws) == pytest.approx(col_width_to_inches(8.43))


@pytest.mark.xfail(reason="Defect 8: formula cells are read as formula text, not values")
def test_defect_8_formula_cells_read_as_cached_values(save_workbook):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 2
    ws["A2"] = "=A1*3"
    path = save_workbook(wb, "formula.xlsx")
    _add_cached_formula_result(path, formula="A1*3", value="6")
    assert load_active_sheet(path)["A2"].value == 6


def _add_cached_formula_result(path, formula, value):
    """Write the result Excel would cache when saving (openpyxl writes none)."""
    sheet = "xl/worksheets/sheet1.xml"
    with zipfile.ZipFile(path) as src:
        parts = {name: src.read(name) for name in src.namelist()}
    old = f"<f>{formula}</f><v></v>".encode()
    assert old in parts[sheet], "unexpected openpyxl output format"
    parts[sheet] = parts[sheet].replace(old, f"<f>{formula}</f><v>{value}</v>".encode())
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as dst:
        for name, data in parts.items():
            dst.writestr(name, data)


def _run_analyze_in_subprocess(folder, filename, encoding):
    """Run analyze() in a child process with `folder` as the current directory."""
    code = "import sys; from intelliconverter.report import analyze; analyze(sys.argv[1])"
    env = {**os.environ, "PYTHONIOENCODING": encoding, "PYTHONPATH": str(REPO_ROOT)}
    return subprocess.run(
        [sys.executable, "-c", code, filename],
        cwd=folder,
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


@pytest.mark.xfail(reason="Defect 9a: a non-Excel file crashes with a raw traceback")
def test_defect_9a_bad_file_gets_a_friendly_error(tmp_path):
    (tmp_path / "not_excel.xlsx").write_text("this is not a spreadsheet", encoding="utf-8")
    result = _run_analyze_in_subprocess(tmp_path, "not_excel.xlsx", encoding="utf-8")
    assert "Traceback" not in result.stderr


@pytest.mark.xfail(reason="Defect 9b: emoji in the output crash non-UTF-8 consoles")
def test_defect_9b_works_on_a_cp1252_console(tmp_path):
    # cp1252 is what Windows uses when output is redirected to a file.
    shutil.copy(SAMPLE_XLSX, tmp_path / "test.xlsx")
    result = _run_analyze_in_subprocess(tmp_path, "test.xlsx", encoding="cp1252")
    assert result.returncode == 0, result.stderr
