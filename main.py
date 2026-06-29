import openpyxl
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

def recommend_layout(total_width_inches):
    """Given total table width, recommend the best page layout."""
    layouts = [
        ("A4 Portrait",     7.7),
        ("A4 Landscape",   10.7),
        ("Legal Landscape", 13.2),
        ("A3 Portrait",    11.7),
    ]
    for name, max_width in layouts:
        if total_width_inches <= max_width:
            scale = (total_width_inches / max_width) * 100
            return name, scale

    return "A3 Landscape", 100  # fallback

def analyze(filepath):
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    total_width = get_total_table_width(ws)
    layout, scale = recommend_layout(total_width)

    print(f"Sheet: {ws.title}")
    print(f"Columns: {ws.max_column}, Rows: {ws.max_row}")
    print(f"Total width: {total_width:.2f} inches")
    print(f"Recommended layout: {layout} (scale: {scale:.0f}%)")

if __name__ == "__main__":
    analyze("test.xlsx")