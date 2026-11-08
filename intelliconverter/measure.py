"""Measuring sheet dimensions from Excel column widths."""
from openpyxl.utils import get_column_letter


def col_width_to_inches(excel_width):
    """Convert Excel column width units to inches."""
    return (excel_width * 7 + 5) / 96

def get_total_table_width(ws):
    """Calculate total width of all columns in inches."""
    total = 0
    default_width = ws.sheet_format.defaultColWidth or 8.43

    for col_num in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_num)
        width = ws.column_dimensions[col_letter].width \
                if col_letter in ws.column_dimensions \
                else default_width
        total += col_width_to_inches(width)

    return total
