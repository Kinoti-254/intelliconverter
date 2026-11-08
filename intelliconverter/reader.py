"""Opening workbooks and turning one sheet into a SheetModel.

load_sheet() checks the file, finds the used range (the cells that actually
hold values), applies the size limits and returns a SheetModel. Every
problem a user can cause is raised as InputError with a plain message.
"""
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException

from intelliconverter.errors import InputError

__all__ = ["ColumnInfo", "InputError", "SheetModel", "load_active_sheet", "load_sheet"]

MAX_FILE_BYTES = 20 * 1024 * 1024
MAX_ROWS = 50_000
MAX_COLUMNS = 100
# Reading takes roughly 20 s and 500 MB per million cells, so the total is
# capped as well as each side.
MAX_CELLS = 1_000_000

DEFAULT_COLUMN_WIDTH = 8.43  # Excel's default, in characters
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
ZIP_SIGNATURE = b"PK\x03\x04"
# Encrypted .xlsx files and old .xls files are both OLE compound files.
OLE_SIGNATURE = bytes.fromhex("d0cf11e0a1b11ae1")


@dataclass(frozen=True)
class ColumnInfo:
    index: int      # 1 for column A
    letter: str
    width: float    # Excel width units (characters)
    hidden: bool


@dataclass
class SheetModel:
    """One sheet's used range: its columns, hidden rows and cell values.

    `cells` holds every row of the used range, hidden ones included, so
    positions match the sheet. Use visible_rows() for what should be shown.
    `formats` holds each cell's Excel number format, in the same positions
    (None means General everywhere).
    """

    title: str
    columns: list[ColumnInfo]
    hidden_rows: frozenset[int]
    cells: list[list]
    uncached_formula_count: int
    formats: list = field(default=None, repr=False)

    @property
    def row_count(self):
        return len(self.cells)

    @property
    def column_count(self):
        return len(self.columns)

    @property
    def visible_columns(self):
        return [column for column in self.columns if not column.hidden]

    @property
    def visible_row_count(self):
        return self.row_count - len(self.hidden_rows)

    def visible_rows(self):
        """Cell values with hidden rows and columns removed."""
        return self._visible(self.cells)

    def visible_formats(self):
        """Number formats in the same shape as visible_rows()."""
        if self.formats is None:
            return [["General"] * len(row) for row in self.visible_rows()]
        return self._visible(self.formats)

    def _visible(self, grid):
        keep = [column.index - 1 for column in self.visible_columns]
        return [
            [row[i] for i in keep]
            for row_number, row in enumerate(grid, start=1)
            if row_number not in self.hidden_rows
        ]


def load_active_sheet(filepath):
    """Open an Excel workbook and return its active worksheet.

    Kept from Phase 0 for existing callers; new code should use load_sheet().
    """
    wb = openpyxl.load_workbook(filepath)
    return wb.active


def load_sheet(filepath, sheet=None):
    """Read one sheet (the active one unless `sheet` names another) into a SheetModel."""
    path = Path(filepath)
    shown = str(filepath)
    _check_file(path, shown)

    # The scan also enforces the size limits, before the slow full load.
    sheet_name, last_row, last_col, formula_cells = _scan_used_range(path, shown, sheet)

    wb = _open_workbook(path, shown, data_only=True)
    try:
        ws = wb[sheet_name]
        columns = _read_columns(ws, last_col)
        hidden_rows = frozenset(
            number for number, dim in ws.row_dimensions.items()
            if dim.hidden and number <= last_row
        )
        cells, formats = [], []
        for row in ws.iter_rows(min_row=1, max_row=last_row, max_col=last_col):
            cells.append([cell.value for cell in row])
            formats.append([cell.number_format for cell in row])
    finally:
        wb.close()

    uncached = sum(1 for row, col in formula_cells if cells[row - 1][col - 1] is None)
    model = SheetModel(sheet_name, columns, hidden_rows, cells, uncached, formats)
    if not model.visible_columns or model.visible_row_count == 0:
        raise InputError(
            f"Sheet '{sheet_name}' has no visible data: all of its rows or columns are hidden."
        )
    return model


def _check_file(path, shown):
    if not path.exists():
        raise InputError(f"File not found: {shown}")
    if not path.is_file():
        raise InputError(f"{shown} is not a file.")

    extension = path.suffix.lower()
    if extension == ".csv":
        raise InputError("CSV files aren't supported yet. Open the file in Excel and save it as .xlsx.")
    if extension == ".xls":
        raise InputError(f"{shown} is in the old .xls format. Open it in Excel and save it as .xlsx.")
    if extension not in EXCEL_EXTENSIONS:
        raise InputError(f"Unsupported file type '{extension}'. IntelliConverter reads Excel .xlsx files.")

    size = path.stat().st_size
    if size > MAX_FILE_BYTES:
        raise InputError(
            f"{shown} is {_megabytes(size)} MB; the limit is {_megabytes(MAX_FILE_BYTES)} MB."
        )

    with path.open("rb") as handle:
        header = handle.read(8)
    if header == OLE_SIGNATURE:
        raise InputError(
            f"{shown} is password-protected or in the old .xls format. "
            "Remove the password or save it as .xlsx, then try again."
        )
    if not header.startswith(ZIP_SIGNATURE):
        raise _not_excel(shown)


def _megabytes(size):
    """20971520 -> '20', 26525287 -> '25.3'."""
    return f"{round(size / (1024 * 1024), 1):g}"


def _not_excel(shown):
    return InputError(
        f"{shown} isn't a valid Excel file. It may be damaged, or a different "
        "kind of file renamed to .xlsx."
    )


def _open_workbook(path, shown, **options):
    try:
        return openpyxl.load_workbook(path, **options)
    except (zipfile.BadZipFile, InvalidFileException, KeyError) as error:
        raise _not_excel(shown) from error
    except OSError as error:
        raise InputError(f"Could not open {shown}: {error.strerror or error}") from error


def _pick_sheet(wb, sheet, shown):
    if sheet is None:
        ws = wb.active
        if ws is None:
            raise InputError(f"{shown} has no sheets.")
    elif sheet in wb.sheetnames:
        ws = wb[sheet]
    else:
        available = ", ".join(wb.sheetnames)
        raise InputError(f"No sheet named '{sheet}' in {shown}. Available sheets: {available}.")

    if not hasattr(ws, "iter_rows"):
        raise InputError(
            f"'{ws.title}' is a chart sheet and has no cells to convert. "
            "Choose a worksheet with --sheet."
        )
    return ws


def _scan_used_range(path, shown, sheet):
    """Find the last row and column holding a value, and where formulas are.

    Uses openpyxl's fast read-only mode. Formatting on empty cells is ignored.
    The used range only grows as the scan goes on, so it stops as soon as a
    size limit is passed instead of reading the rest of a huge sheet.
    """
    wb = _open_workbook(path, shown, read_only=True)
    try:
        ws = _pick_sheet(wb, sheet, shown)
        # Some programs write a wrong size into the file; without this,
        # read-only mode trusts it and can silently skip cells.
        ws.reset_dimensions()

        last_row = last_col = 0
        formula_cells = []
        for row_number, row in enumerate(ws.iter_rows(values_only=True), start=1):
            for col_number, value in enumerate(row, start=1):
                if value is None or value == "":
                    continue
                if row_number > last_row or col_number > last_col:
                    last_row = row_number
                    last_col = max(last_col, col_number)
                    _check_limits(last_row, last_col)
                if isinstance(value, str) and value.startswith("="):
                    formula_cells.append((row_number, col_number))

        if last_row == 0:
            raise InputError(f"Sheet '{ws.title}' is empty.")
        return ws.title, last_row, last_col, formula_cells
    finally:
        wb.close()


def _check_limits(rows, columns):
    """Called during the scan with the used range so far, hence 'at least'."""
    if rows > MAX_ROWS:
        raise InputError(f"The sheet has at least {rows:,} rows; the limit is {MAX_ROWS:,}.")
    if columns > MAX_COLUMNS:
        raise InputError(f"The sheet has at least {columns:,} columns; the limit is {MAX_COLUMNS:,}.")
    cells = rows * columns
    if cells > MAX_CELLS:
        raise InputError(
            f"The sheet has at least {cells:,} cells ({rows:,} rows x {columns:,} columns); "
            f"the limit is {MAX_CELLS:,}."
        )


def _read_columns(ws, last_col):
    """Width and hidden flag for each column up to last_col.

    openpyxl stores a run of identical columns (such as B:E) as one entry
    keyed by its first letter, so each entry is expanded over its range.
    """
    default_width = ws.sheet_format.defaultColWidth or DEFAULT_COLUMN_WIDTH
    widths = {}
    hidden = set()
    for dim in ws.column_dimensions.values():
        for index in range(dim.min, min(dim.max, last_col) + 1):
            widths[index] = dim.width
            if dim.hidden:
                hidden.add(index)

    return [
        ColumnInfo(
            index=index,
            letter=get_column_letter(index),
            width=widths.get(index, default_width),
            hidden=index in hidden,
        )
        for index in range(1, last_col + 1)
    ]
