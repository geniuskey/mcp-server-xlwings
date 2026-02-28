"""MCP server for Excel automation via xlwings COM."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from .excel import ExcelError, ExcelHandler

mcp = FastMCP(
    "mcp-server-xlwings",
    instructions=(
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


# ================================================================== #
#  Tool 1: get_active_workbook
# ================================================================== #


@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Active Workbook",
        readOnlyHint=True,
    ),
)
def get_active_workbook() -> dict:
    """Get the currently active workbook info including sheets,
    active sheet, and current selection address with its data."""
    return _call(_handler.get_active_workbook)


# ================================================================== #
#  Tool 2: manage_workbooks
# ================================================================== #


@mcp.tool(
    annotations=ToolAnnotations(
        title="Manage Workbooks",
        destructiveHint=True,
    ),
)
def manage_workbooks(
    action: str,
    workbook: str | None = None,
    filepath: str | None = None,
    read_only: bool = True,
    save: bool = False,
) -> dict:
    """Manage Excel workbooks: list, open, save, close, or recalculate.

    Args:
        action: One of 'list', 'open', 'save', 'close', 'recalculate'.
        workbook: Workbook name or path. Defaults to active workbook.
        filepath: For 'open': file path (use 'new' for blank). For 'save': Save As path.
        read_only: For 'open': open in read-only mode.
        save: For 'close': save before closing.
    """
    return _call(
        _handler.manage_workbooks, action, workbook, filepath, read_only, save
    )


# ================================================================== #
#  Tool 3: read_data
# ================================================================== #


@mcp.tool(
    annotations=ToolAnnotations(
        title="Read Data",
        readOnlyHint=True,
    ),
)
def read_data(
    workbook: str | None = None,
    sheet: str | None = None,
    cell_range: str | None = None,
    headers: bool = True,
    detail: bool = False,
) -> dict:
    """Read data from an Excel range. Reads the entire used range when cell_range is omitted.
    Set detail=True on a single cell to get formula, type, and formatting info.

    Args:
        workbook: Workbook name or path. Defaults to active workbook.
        sheet: Sheet name. Defaults to active sheet.
        cell_range: Range like 'A1:D10' or cell like 'B5'. Reads used range if omitted.
        headers: Treat first row as column headers.
        detail: For single cells, include formula, type, number format, and font info.
    """
    return _call(
        _handler.read_data, workbook, sheet, cell_range, headers, detail
    )


# ================================================================== #
#  Tool 4: write_data
# ================================================================== #


@mcp.tool(
    annotations=ToolAnnotations(
        title="Write Data",
        destructiveHint=True,
    ),
)
def write_data(
    start_cell: str,
    data: list[list] | None = None,
    formula: str | None = None,
    workbook: str | None = None,
    sheet: str | None = None,
) -> dict:
    """Write data or a formula to Excel cells.
    Provide 'data' for a 2D array, or 'formula' for a single-cell formula.

    Args:
        start_cell: Top-left cell (e.g. 'A1').
        data: 2D list of values. Mutually exclusive with formula.
        formula: Excel formula like '=SUM(A1:A10)'. Mutually exclusive with data.
        workbook: Workbook name or path. Defaults to active workbook.
        sheet: Sheet name. Defaults to active sheet.
    """
    return _call(
        _handler.write_data, start_cell, data, formula, workbook, sheet
    )


# ================================================================== #
#  Tool 5: manage_sheets
# ================================================================== #


@mcp.tool(
    annotations=ToolAnnotations(
        title="Manage Sheets",
        destructiveHint=True,
    ),
)
def manage_sheets(
    action: str,
    workbook: str | None = None,
    sheet: str | None = None,
    new_name: str | None = None,
    position: int = 1,
    count: int = 1,
) -> dict:
    """Manage sheets and structure.

    Args:
        action: One of 'list', 'add', 'delete', 'rename', 'copy', 'activate',
                'insert_rows', 'delete_rows', 'insert_columns', 'delete_columns'.
        workbook: Workbook name or path. Defaults to active workbook.
        sheet: Target sheet name (required for delete/rename/copy/activate).
        new_name: New name (for rename; optional for add/copy).
        position: Row/column number (1-based) for insert/delete actions.
        count: Number of rows/columns to insert or delete.
    """
    return _call(
        _handler.manage_sheets,
        action, workbook, sheet, new_name, position, count,
    )


# ================================================================== #
#  Tool 6: find_replace
# ================================================================== #


@mcp.tool(
    annotations=ToolAnnotations(
        title="Find and Replace",
        destructiveHint=False,
    ),
)
def find_replace(
    find: str,
    workbook: str | None = None,
    sheet: str | None = None,
    replace: str | None = None,
    match_case: bool = False,
) -> dict:
    """Search for text in a sheet, optionally replacing it.

    Args:
        find: Text to search for.
        workbook: Workbook name or path. Defaults to active workbook.
        sheet: Sheet name. Defaults to active sheet.
        replace: Replacement text. If omitted, search only.
        match_case: Case-sensitive matching.
    """
    return _call(
        _handler.find_replace, find, workbook, sheet, replace, match_case
    )


# ================================================================== #
#  Tool 7: format_range
# ================================================================== #


@mcp.tool(
    annotations=ToolAnnotations(
        title="Format Range",
        destructiveHint=True,
    ),
)
def format_range(
    cell_range: str,
    workbook: str | None = None,
    sheet: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    font_size: int | None = None,
    font_color: str | None = None,
    bg_color: str | None = None,
    number_format: str | None = None,
    alignment: str | None = None,
    wrap_text: bool | None = None,
    border: bool | None = None,
) -> dict:
    """Apply formatting to a cell range.

    Args:
        cell_range: Range like 'A1:D10'.
        workbook: Workbook name or path. Defaults to active workbook.
        sheet: Sheet name. Defaults to active sheet.
        bold: Set bold.
        italic: Set italic.
        underline: Set underline.
        font_size: Font size in points.
        font_color: Hex colour like '#FF0000'.
        bg_color: Background hex colour like '#FFFF00'.
        number_format: Excel format like '#,##0.00'.
        alignment: 'left', 'center', 'right', 'justify'.
        wrap_text: Enable text wrapping.
        border: Apply thin borders.
    """
    return _call(
        _handler.format_range,
        cell_range, workbook, sheet,
        bold, italic, underline, font_size, font_color,
        bg_color, number_format, alignment, wrap_text, border,
    )


# ================================================================== #
#  Tool 8: run_macro
# ================================================================== #


@mcp.tool(
    annotations=ToolAnnotations(
        title="Run Macro",
        destructiveHint=True,
    ),
)
def run_macro(
    macro_name: str,
    workbook: str | None = None,
    args: list | None = None,
) -> dict:
    """Run a VBA macro in Excel and return its result.

    Args:
        macro_name: Macro name (e.g. 'MyMacro' or 'Module1.MyMacro').
        workbook: Workbook name. If omitted, Excel resolves globally.
        args: Optional arguments to pass to the macro.
    """
    return _call(_handler.run_macro, macro_name, workbook, args)
