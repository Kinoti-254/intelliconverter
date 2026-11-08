"""Tests for the command line: python main.py [file] [--sheet NAME].

Includes the regression tests for defects 9a and 9b, fixed in Phase 1.
"""
import os
import shutil
import subprocess
import sys

import openpyxl
import pytest

from intelliconverter import cli
from tests.conftest import GOLDEN_DIR, REPO_ROOT, SAMPLE_XLSX, read_pdf

GOLDEN_REPORT = (GOLDEN_DIR / "test_xlsx_report.txt").read_text(encoding="utf-8")


def run_main(folder, *args, encoding="utf-8"):
    """Run `python main.py ARGS` in a child process with `folder` as the current directory."""
    env = {**os.environ, "PYTHONIOENCODING": encoding}
    return subprocess.run(
        [sys.executable, str(REPO_ROOT / "main.py"), *args],
        cwd=folder,
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


@pytest.fixture
def two_sheet_book(workdir):
    wb = openpyxl.Workbook()
    wb.active.title = "Summary"
    wb.active.append(["Total", 10])
    detail = wb.create_sheet("Detail")
    detail.append(["Item", "Qty"])
    detail.append(["Widget", 7])
    wb.save(workdir / "book.xlsx")
    return workdir / "book.xlsx"


# --- Normal use -------------------------------------------------------------

def test_no_arguments_analyses_test_xlsx(workdir, capsys):
    assert cli.main([]) == 0
    assert capsys.readouterr().out == GOLDEN_REPORT
    assert (workdir / "output.pdf").is_file()


def test_file_argument(two_sheet_book, capsys):
    assert cli.main([str(two_sheet_book)]) == 0
    out = capsys.readouterr().out
    assert "Sheet   : Summary" in out
    assert "Total" in read_pdf(two_sheet_book.parent / "output.pdf")["text"]


def test_sheet_option(two_sheet_book, capsys):
    assert cli.main([str(two_sheet_book), "--sheet", "Detail"]) == 0
    assert "Sheet   : Detail" in capsys.readouterr().out
    text = read_pdf(two_sheet_book.parent / "output.pdf")["text"]
    assert "Widget" in text and "Total" not in text


def test_help(capsys):
    with pytest.raises(SystemExit) as excinfo:
        cli.main(["--help"])
    assert excinfo.value.code == 0
    out = capsys.readouterr().out
    assert "--sheet" in out and "file" in out


def test_running_main_py_without_arguments_matches_original(tmp_path):
    shutil.copy(SAMPLE_XLSX, tmp_path / "test.xlsx")
    result = run_main(tmp_path)
    assert result.returncode == 0, result.stderr
    assert result.stdout.replace("\r\n", "\n") == GOLDEN_REPORT
    assert (tmp_path / "output.pdf").is_file()


# --- Errors the user can hit ------------------------------------------------

def test_missing_file_is_a_friendly_error(workdir, capsys):
    assert cli.main(["nope.xlsx"]) == 1
    captured = capsys.readouterr()
    assert captured.err == "Error: File not found: nope.xlsx\n"
    assert captured.out == ""


def test_unknown_sheet_is_a_friendly_error(two_sheet_book, capsys):
    assert cli.main([str(two_sheet_book), "--sheet", "Nope"]) == 1
    assert "No sheet named 'Nope'" in capsys.readouterr().err


def test_unwritable_output_pdf_is_a_friendly_error(workdir, capsys):
    # A folder named output.pdf stands in for a PDF that is open in another
    # program (on Windows, an open PDF cannot be overwritten).
    (workdir / "output.pdf").mkdir()
    assert cli.main([]) == 1
    err = capsys.readouterr().err
    assert err.startswith("Error: Could not write output.pdf.")
    assert "open in another program" in err


def test_uncached_formulas_print_a_warning(workdir, capsys):
    wb = openpyxl.Workbook()
    wb.active["A1"] = 2
    wb.active["A2"] = "=A1*3"
    wb.save(workdir / "formulas.xlsx")
    assert cli.main(["formulas.xlsx"]) == 0
    assert "Warning : 1 formula cell has no saved value" in capsys.readouterr().out


def test_defect_9a_bad_file_gets_a_friendly_error(tmp_path):
    (tmp_path / "not_excel.xlsx").write_text("this is not a spreadsheet", encoding="utf-8")
    result = run_main(tmp_path, "not_excel.xlsx")
    assert result.returncode == 1
    assert "Traceback" not in result.stderr
    assert result.stderr.startswith("Error: not_excel.xlsx isn't a valid Excel file")


def test_defect_9b_works_on_a_cp1252_console(tmp_path):
    # cp1252 is what Windows uses when output is redirected to a file.
    shutil.copy(SAMPLE_XLSX, tmp_path / "test.xlsx")
    result = run_main(tmp_path, encoding="cp1252")
    assert result.returncode == 0, result.stderr
    assert "Best layout : A4 Portrait" in result.stdout
    assert (tmp_path / "output.pdf").is_file()
