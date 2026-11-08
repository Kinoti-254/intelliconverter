"""Characterisation tests for the console report, the reader and main.py."""
import os
import shutil
import subprocess
import sys

import openpyxl
from openpyxl.utils import get_column_letter

import main
from intelliconverter import layout, measure, report
from intelliconverter.reader import load_active_sheet
from tests.conftest import GOLDEN_DIR, REPO_ROOT, SAMPLE_XLSX

# Console output of the original code for `python main.py` on test.xlsx,
# recorded before the restructure. It includes the PDF confirmation lines.
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


def test_reader_returns_the_active_sheet(save_workbook):
    wb = openpyxl.Workbook()
    wb.active.title = "First"
    wb.create_sheet("Second")
    wb.active = 1
    assert load_active_sheet(save_workbook(wb)).title == "Second"


def test_main_reexports_the_original_names():
    assert main.analyze is report.analyze
    assert main.score_layout is layout.score_layout
    assert main.LAYOUTS is layout.LAYOUTS
    assert main.col_width_to_inches is measure.col_width_to_inches
    assert main.get_total_table_width is measure.get_total_table_width


def test_running_main_py_prints_original_output_and_writes_pdf(tmp_path):
    shutil.copy(SAMPLE_XLSX, tmp_path / "test.xlsx")
    # UTF-8 output is forced because the report contains emoji; see the
    # known defect about non-UTF-8 consoles.
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    result = subprocess.run(
        [sys.executable, str(REPO_ROOT / "main.py")],
        cwd=tmp_path,
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    assert result.returncode == 0, result.stderr
    assert result.stdout.replace("\r\n", "\n") == GOLDEN_REPORT
    assert (tmp_path / "output.pdf").is_file()
