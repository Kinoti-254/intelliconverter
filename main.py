from openpyxl import load_workbook

wb = load_workbook("test.xlsx")
ws = wb.active

print(ws)
print(ws.dimensions)

for col_letter, dim in ws.column_dimensions.items():
    print(col_letter, dim.width)

print(ws.sheet_format.defaultColWidth)

def col_width_to_inches(excel_width):
    return (excel_width * 7 + 5) / 96

print(col_width_to_inches(12.63))
print(col_width_to_inches(25.63))

from openpyxl.utils import get_column_letter

for col_num in range(1, ws.max_column + 1):
    col_letter = get_column_letter(col_num)
    if col_letter in ws.column_dimensions:
        width = ws.column_dimensions[col_letter].width
    else:
        width = ws.sheet_format.defaultColWidth
    print(col_letter, "width =", width)