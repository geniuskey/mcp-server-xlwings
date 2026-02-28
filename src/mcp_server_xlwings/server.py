"""MCP server for Excel automation via xlwings COM."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from .excel import ExcelError, ExcelHandler

mcp = FastMCP(
    "mcp-server-xlwings",
    description=(
        "MCP server for Excel automation via xlwings COM. "
        "Works with DRM-protected files by controlling the live Excel process."
    ),
)

_handler = ExcelHandler()


def _call(fn, *args, **kwargs):
    """Invoke an ExcelHandler method with uniform error handling."""
    try:
        return fn(*args, **kwargs)
    except ExcelError:
        raise
    except Exception as exc:
        raise ExcelError(f"Excel COM error: {exc}") from exc


# ------------------------------------------------------------------ #
#  Tools
# ------------------------------------------------------------------ #


@mcp.tool()
def list_open_workbooks() -> list[dict]:
    """List all currently open Excel workbooks with their sheet names."""
    return _call(_handler.list_open_workbooks)


@mcp.tool()
def open_workbook(filepath: str, read_only: bool = True) -> dict:
    """Open an Excel file or create a new workbook.

    Args:
        filepath: File path to open. Use "new" to create a blank workbook.
        read_only: Open in read-only mode (default True).
    """
    return _call(_handler.open_workbook, filepath, read_only)


@mcp.tool()
def read_range(
    workbook: str,
    sheet: str | None = None,
    cell_range: str | None = None,
    headers: bool = True,
) -> dict:
    """Read data from a specified range in an Excel workbook.

    Args:
        workbook: Workbook name (e.g. 'report.xlsx') or full file path.
        sheet: Sheet name. Uses the active sheet when omitted.
        cell_range: Range like 'A1:D10'. Reads the entire used range when omitted.
        headers: Treat the first row as column headers.
    """
    return _call(_handler.read_range, workbook, sheet, cell_range, headers)


@mcp.tool()
def write_range(
    workbook: str,
    start_cell: str,
    data: list[list],
    sheet: str | None = None,
) -> dict:
    """Write a 2D array of data to the specified location.

    Args:
        workbook: Workbook name or full file path.
        start_cell: Top-left cell to start writing (e.g. 'A1').
        data: 2D list of values to write.
        sheet: Sheet name. Uses the active sheet when omitted.
    """
    return _call(_handler.write_range, workbook, sheet, start_cell, data)


@mcp.tool()
def read_cell_info(
    workbook: str,
    cell: str,
    sheet: str | None = None,
) -> dict:
    """Get detailed information about a single cell (value, formula, format, type).

    Args:
        workbook: Workbook name or full file path.
        cell: Cell address like 'B5'.
        sheet: Sheet name. Uses the active sheet when omitted.
    """
    return _call(_handler.read_cell_info, workbook, sheet, cell)


@mcp.tool()
def set_formula(
    workbook: str,
    cell: str,
    formula: str,
    sheet: str | None = None,
) -> dict:
    """Set a formula in a cell and return the calculated value.

    Args:
        workbook: Workbook name or full file path.
        cell: Target cell like 'C10'.
        formula: Excel formula like '=SUM(A1:A10)'.
        sheet: Sheet name. Uses the active sheet when omitted.
    """
    return _call(_handler.set_formula, workbook, sheet, cell, formula)


@mcp.tool()
def manage_sheets(
    workbook: str,
    action: str,
    sheet: str | None = None,
    new_name: str | None = None,
) -> dict:
    """Manage sheets: list, add, delete, rename, or copy.

    Args:
        workbook: Workbook name or full file path.
        action: One of 'list', 'add', 'delete', 'rename', 'copy'.
        sheet: Target sheet name (required for delete/rename/copy).
        new_name: New sheet name (required for rename; optional for add/copy).
    """
    return _call(_handler.manage_sheets, workbook, action, sheet, new_name)


@mcp.tool()
def find_replace(
    workbook: str,
    find: str,
    sheet: str | None = None,
    replace: str | None = None,
    match_case: bool = False,
) -> dict:
    """Search for text in a sheet, optionally replacing it.

    Args:
        workbook: Workbook name or full file path.
        find: Text to search for.
        sheet: Sheet name. Uses the active sheet when omitted.
        replace: Replacement text. If omitted, search only.
        match_case: Case-sensitive matching.
    """
    return _call(
        _handler.find_replace, workbook, sheet, find, replace, match_case
    )


@mcp.tool()
def format_range(
    workbook: str,
    cell_range: str,
    sheet: str | None = None,
    bold: bool | None = None,
    font_size: int | None = None,
    font_color: str | None = None,
    bg_color: str | None = None,
    number_format: str | None = None,
    border: bool | None = None,
) -> dict:
    """Apply formatting to a cell range.

    Args:
        workbook: Workbook name or full file path.
        cell_range: Range like 'A1:D10'.
        sheet: Sheet name. Uses the active sheet when omitted.
        bold: Set bold.
        font_size: Font size in points.
        font_color: Font colour as hex string like '#FF0000'.
        bg_color: Background colour as hex string like '#FFFF00'.
        number_format: Excel number format like '#,##0.00'.
        border: Apply thin borders to all edges.
    """
    return _call(
        _handler.format_range,
        workbook,
        sheet,
        cell_range,
        bold,
        font_size,
        font_color,
        bg_color,
        number_format,
        border,
    )


@mcp.tool()
def save_workbook(
    workbook: str,
    filepath: str | None = None,
) -> dict:
    """Save the workbook. Optionally save to a new path (Save As).

    Args:
        workbook: Workbook name or full file path.
        filepath: New file path for Save As. Omit to overwrite the current file.
    """
    return _call(_handler.save_workbook, workbook, filepath)
