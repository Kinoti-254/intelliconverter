"""Opening workbooks and turning sheets into SheetModels.

load_sheet() reads one sheet; load_workbook_sheets() reads several (by
default every visible worksheet). Both check the file, find each sheet's
used range (the cells that actually hold values), apply the size limits and
build SheetModels. The file is opened once in openpyxl's fast read-only mode
(to scan) and once in full mode (to read), however many sheets are read.
Every problem a user can cause is raised as InputError with a plain message.
"""
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException

from intelliconverter.errors import InputError

__all__ = ["ColumnInfo", "InputError", "SheetModel", "WorkbookSheets",
           "load_active_sheet", "load_sheet", "load_workbook_sheets"]

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


@dataclass
class WorkbookSheets:
    """Sheets read from one workbook, in conversion order."""

    models: list    # SheetModel per converted sheet
    skipped: list   # (sheet name, reason) for sheets left out, in workbook order


def load_active_sheet(filepath):
    """Open an Excel workbook and return its active worksheet.

    Kept from Phase 0 for existing callers; new code should use load_sheet().
    """
    wb = openpyxl.load_workbook(filepath)
    return wb.active


def load_sheet(filepath, sheet=None):
    """Read one sheet (the active one unless `sheet` names another) into a SheetModel."""
    return _load(filepath, [sheet]).models[0]


def load_workbook_sheets(filepath, names=None):
    """Read several sheets. Returns WorkbookSheets.

    names=None reads every visible worksheet, in workbook order, and skips
    (with a reason) hidden sheets, chart sheets, empty sheets and sheets
    whose rows or columns are all hidden. Named sheets are read in the
    order given (repeats once); one that can't be read is an error, even if
    it is hidden.
    """
    if names is None:
        return _load(filepath, None)
    return _load(filepath, list(dict.fromkeys(names)))


def _load(filepath, names):
    """names: a list of sheet names (None in it means the active sheet), or None for all."""
    path = Path(filepath)
    shown = str(filepath)
    _check_file(path, shown)
    strict = names is not None

    scans, skipped, order = _scan_sheets(path, shown, names)
    models = []
    if scans:
        wb = _open_workbook(path, shown, data_only=True)
        try:
            for name, last_row, last_col, formula_cells in scans:
                model = _build_model(wb[name], last_row, last_col, formula_cells)
                if model.visible_columns and model.visible_row_count:
                    models.append(model)
                elif strict:
                    raise InputError(
                        f"Sheet '{name}' has no visible data: all of its rows or columns are hidden."
                    )
                else:
                    skipped.append((name, "nothing visible"))
        finally:
            wb.close()
    skipped.sort(key=lambda item: order[item[0]])
    return WorkbookSheets(models, skipped)


def _scan_sheets(path, shown, names):
    """Scan each sheet to convert in read-only mode, enforcing the limits.

    Returns (scans, skipped, order): one (name, last_row, last_col,
    formula_cells) per sheet to read, the sheets skipped so far, and each
    sheet's position in the workbook.
    """
    wb = _open_workbook(path, shown, read_only=True)
    try:
        order = {name: index for index, name in enumerate(wb.sheetnames)}
        scans, skipped = [], []
        total_cells = 0
        for target in (wb.sheetnames if names is None else names):
            if names is None:
                ws = wb[target]
                reason = _skip_reason(ws)
                if reason:
                    skipped.append((target, reason))
                    continue
            else:
                ws = _pick_sheet(wb, target, shown)
            scan = _scan_used_range(ws, total_cells)
            if scan is None:
                if names is not None:
                    raise InputError(f"Sheet '{ws.title}' is empty.")
                skipped.append((ws.title, "empty"))
                continue
            last_row, last_col, formula_cells = scan
            total_cells += last_row * last_col
            scans.append((ws.title, last_row, last_col, formula_cells))
        return scans, skipped, order
    finally:
        wb.close()


def _skip_reason(ws):
    if not hasattr(ws, "iter_rows"):
        return "chart sheet"
    if ws.sheet_state != "visible":
        return "hidden sheet"
    return None


def _build_model(ws, last_row, last_col, formula_cells):
    columns = _read_columns(ws, last_col)
    hidden_rows = frozenset(
        number for number, dim in ws.row_dimensions.items()
        if dim.hidden and number <= last_row
    )
    cells, formats = [], []
    for row in ws.iter_rows(min_row=1, max_row=last_row, max_col=last_col):
        cells.append([cell.value for cell in row])
        formats.append([cell.number_format for cell in row])
    uncached = sum(1 for row, col in formula_cells if cells[row - 1][col - 1] is None)
    return SheetModel(ws.title, columns, hidden_rows, cells, uncached, formats)


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


def _scan_used_range(ws, cells_before):
    """Find the last row and column holding a value, and where formulas are.

    Returns (last_row, last_col, formula_cells), or None for an empty sheet.
    Formatting on empty cells is ignored. The used range only grows as the
    scan goes on, so it stops as soon as a size limit is passed (counting
    `cells_before` from sheets already scanned) instead of reading the rest.
    """
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
                _check_limits(ws.title, last_row, last_col, cells_before)
            if isinstance(value, str) and value.startswith("="):
                formula_cells.append((row_number, col_number))

    if last_row == 0:
        return None
    return last_row, last_col, formula_cells


def _check_limits(name, rows, columns, cells_before=0):
    """Called during the scan with the used range so far, hence 'at least'."""
    if rows > MAX_ROWS:
        raise InputError(f"Sheet '{name}' has at least {rows:,} rows; the limit is {MAX_ROWS:,}.")
    if columns > MAX_COLUMNS:
        raise InputError(
            f"Sheet '{name}' has at least {columns:,} columns; the limit is {MAX_COLUMNS:,}."
        )
    cells = rows * columns
    if cells > MAX_CELLS:
        raise InputError(
            f"Sheet '{name}' has at least {cells:,} cells ({rows:,} rows x {columns:,} columns); "
            f"the limit is {MAX_CELLS:,}."
        )
    if cells_before + cells > MAX_CELLS:
        raise InputError(
            f"The selected sheets have at least {cells_before + cells:,} cells in total; "
            f"the limit is {MAX_CELLS:,}. Convert fewer sheets with --sheet."
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
