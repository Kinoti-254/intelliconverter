"""Shared test helpers.

Test workbooks are generated in code, so each test shows exactly what its
file contains. Anything that writes files runs in a temporary folder, so
the tests never touch test.xlsx or output.pdf in the project folder.
"""
import shutil
from pathlib import Path

import openpyxl
import pytest
import reportlab.rl_config
from pypdf import PdfReader

import intelliconverter.render

REPO_ROOT = Path(__file__).resolve().parent.parent
SAMPLE_XLSX = REPO_ROOT / "test.xlsx"
GOLDEN_DIR = Path(__file__).resolve().parent / "golden"


@pytest.fixture
def save_workbook(tmp_path):
    """Return a function that saves a workbook to a temp file and returns its path."""

    def _save(wb, name="book.xlsx"):
        path = tmp_path / name
        wb.save(path)
        return path

    return _save


@pytest.fixture
def reload_sheet(save_workbook):
    """Save a workbook, reopen it, and return the active sheet.

    Reloading matters: some defects only appear after openpyxl reads
    a file back from disk (for example, grouped column ranges).
    """

    def _reload(wb):
        path = save_workbook(wb)
        return openpyxl.load_workbook(path).active

    return _reload


@pytest.fixture
def workdir(tmp_path, monkeypatch):
    """A temp folder set as the current directory, holding a copy of test.xlsx.

    analyze() reads test.xlsx and writes output.pdf in the current directory.
    """
    shutil.copy(SAMPLE_XLSX, tmp_path / "test.xlsx")
    monkeypatch.chdir(tmp_path)
    return tmp_path


@pytest.fixture
def deterministic_pdfs(monkeypatch):
    """Switch on ReportLab's invariant mode (fixed dates and IDs) for this test only."""
    monkeypatch.setattr(reportlab.rl_config, "invariant", 1)


@pytest.fixture
def captured_tables(monkeypatch):
    """Record every ReportLab Table the renderer creates (one per page).

    Each entry is a dict with the table's data, column widths and row heights.
    The real Table is still built, so the PDF is unaffected.
    """
    tables = []
    real_table = intelliconverter.render.Table

    def recording_table(data, colWidths=None, rowHeights=None, **kwargs):
        tables.append({"data": data, "col_widths": colWidths, "row_heights": rowHeights})
        return real_table(data, colWidths=colWidths, rowHeights=rowHeights, **kwargs)

    monkeypatch.setattr(intelliconverter.render, "Table", recording_table)
    return tables


def read_pdf(path):
    """Return the page count, page sizes in points, and text of a PDF (whole and per page)."""
    reader = PdfReader(str(path))
    sizes = [(float(p.mediabox.width), float(p.mediabox.height)) for p in reader.pages]
    page_texts = [page.extract_text() for page in reader.pages]
    return {"pages": len(reader.pages), "sizes": sizes, "text": "\n".join(page_texts),
            "page_texts": page_texts}
