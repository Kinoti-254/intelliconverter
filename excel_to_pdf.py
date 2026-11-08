"""Backwards-compatible wrapper for code written before Phase 2b.

build_pdf() now chooses the layout itself (intelliconverter.render). The old
layout_name and scale_percent arguments are still accepted but ignored, with
a DeprecationWarning. The old helpers below are kept so existing imports
keep working; new code should use intelliconverter.render.
"""
import warnings

from reportlab.lib.pagesizes import A3, A4, LEGAL, landscape, portrait

from intelliconverter import render

__all__ = ["PAGE_SIZES", "build_pdf", "extract_sheet_data", "get_page_size"]

# Old layout names and their ReportLab page sizes (no longer used by the engine)
PAGE_SIZES = {
    "A4 Portrait":      portrait(A4),
    "A4 Landscape":     landscape(A4),
    "A3 Portrait":      portrait(A3),
    "A3 Landscape":     landscape(A3),
    "Legal Landscape":  landscape(LEGAL),
}


def get_page_size(layout_name):
    """Return the ReportLab page size for an old layout name."""
    return PAGE_SIZES.get(layout_name, landscape(A4))


def extract_sheet_data(ws):
    """Extract all cell values from an openpyxl worksheet as a 2D list of strings."""
    return [[str(cell) if cell is not None else "" for cell in row]
            for row in ws.iter_rows(values_only=True)]


def build_pdf(filepath, output_path, layout_name=None, scale_percent=None, sheet=None):
    """Generate a PDF from an Excel file; the layout is chosen automatically."""
    if layout_name is not None or scale_percent is not None:
        warnings.warn(
            "build_pdf() now chooses the layout itself; layout_name and "
            "scale_percent are ignored.",
            DeprecationWarning,
            stacklevel=2,
        )
    return render.build_pdf(filepath, output_path, sheet=sheet)
