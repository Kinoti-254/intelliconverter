"""Measuring sheet dimensions from Excel column widths.

The layout engine sizes columns from their content, so this module is not
used for layout today. It is kept for the planned "preserve Excel widths"
mode.
"""


def col_width_to_inches(excel_width):
    """Convert Excel column width units to inches."""
    return (excel_width * 7 + 5) / 96


def get_total_table_width(sheet):
    """Calculate the total width of a SheetModel's visible columns in inches.

    The reader has already applied the default width, expanded column
    ranges and limited the columns to the used range.
    """
    total = 0
    for column in sheet.visible_columns:
        total += col_width_to_inches(column.width)
    return total
