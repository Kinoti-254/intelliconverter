"""Opening workbooks."""
import openpyxl


def load_active_sheet(filepath):
    """Open an Excel workbook and return its active worksheet."""
    wb = openpyxl.load_workbook(filepath)
    return wb.active
