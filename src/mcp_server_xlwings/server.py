"""MCP server for Excel automation via xlwings COM."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

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


# ================================================================== #
#  Active Excel tools  (xlwings-exclusive — openpyxl cannot do these)
# ================================================================== #


@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Active Workbook",
        readOnlyHint=True,
    ),
)
def get_active_workbook() -> dict:
    """Get information about the currently active workbook in Excel,
    including the active sheet and current selection.
    Useful to understand what the user is looking at right now."""
    return _call(_handler.get_active_workbook)


@mcp.tool(
    annotations=ToolAnnotations(
        title="Read Selection",
        readOnlyHint=True,
    ),
)
def read_selection() -> dict:
    """Read the data from the cells currently selected by the user in Excel.
    Returns the values of whatever range the user has highlighted."""
    return _call(_handler.read_selection)


@mcp.tool(
    annotations=ToolAnnotations(
        title="Activate Sheet",
        readOnlyHint=False,
    ),
)
def activate_sheet(workbook: str, sheet: str) -> dict:
    """Switch the active sheet in a workbook so it becomes visible to the user.

    Args:
        workbook: Workbook name or full file path.
        sheet: Sheet name to activate.
    """
    return _call(_handler.activate_sheet, workbook, sheet)


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
    This is impossible with file-based libraries like openpyxl.

    Args:
        macro_name: Macro name (e.g. 'MyMacro' or 'Module1.MyMacro').
        workbook: Workbook name. If omitted, Excel resolves the macro globally.
        args: Optional list of arguments to pass to the macro.
    """
    return _call(_handler.run_macro, macro_name, workbook, args)


@mcp.tool(
    annotations=ToolAnnotations(
        title="Close Workbook",
        destructiveHint=True,
    ),
)
def close_workbook(workbook: str, save: bool = False) -> dict:
    """Close a workbook, optionally saving it first.

    Args:
        workbook: Workbook name or full file path.
        save: Save the workbook before closing.
    """
    return _call(_handler.close_workbook, workbook, save)


@mcp.tool(
    annotations=ToolAnnotations(
        title="Recalculate",
        readOnlyHint=False,
    ),
)
def recalculate(workbook: str | None = None) -> dict:
    """Force Excel to recalculate all formulas.

    Args:
        workbook: Workbook name. If omitted, recalculates all open workbooks.
    """
    return _call(_handler.recalculate, workbook)


# ================================================================== #
#  Core workbook tools
# ================================================================== #


@mcp.tool(
    annotations=ToolAnnotations(
        title="List Open Workbooks",
        readOnlyHint=True,
    ),
)
def list_open_workbooks() -> list[dict]:
    """List all currently open Excel workbooks with their sheet names."""
    return _call(_handler.list_open_workbooks)


@mcp.tool(
    annotations=ToolAnnotations(
        title="Open Workbook",
        readOnlyHint=False,
    ),
)
def open_workbook(filepath: str, read_only: bool = True) -> dict:
    """Open an Excel file or create a new workbook.

    Args:
        filepath: File path to open. Use "new" to create a blank workbook.
        read_only: Open in read-only mode (default True).
    """
    return _call(_handler.open_workbook, filepath, read_only)


@mcp.tool(
    annotations=ToolAnnotations(
        title="Save Workbook",
        destructiveHint=True,
    ),
)
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


# ================================================================== #
#  Read / Write tools
# ================================================================== #


@mcp.tool(
    annotations=ToolAnnotations(
        title="Read Range",
        readOnlyHint=True,
    ),
)
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


@mcp.tool(
    annotations=ToolAnnotations(
        title="Write Range",
        destructiveHint=True,
    ),
)
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


@mcp.tool(
    annotations=ToolAnnotations(
        title="Read Cell Info",
        readOnlyHint=True,
    ),
)
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


@mcp.tool(
    annotations=ToolAnnotations(
        title="Set Formula",
        destructiveHint=True,
    ),
)
def set_formula(
    workbook: str,
    cell: str,
    formula: str,
    sheet: str | None = None,
) -> dict:
    """Set a formula in a cell and return the calculated value.
    Unlike openpyxl, xlwings returns the live calculated result immediately.

    Args:
        workbook: Workbook name or full file path.
        cell: Target cell like 'C10'.
        formula: Excel formula like '=SUM(A1:A10)'.
        sheet: Sheet name. Uses the active sheet when omitted.
    """
    return _call(_handler.set_formula, workbook, sheet, cell, formula)


@mcp.tool(
    annotations=ToolAnnotations(
        title="Find and Replace",
        destructiveHint=False,
    ),
)
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


# ================================================================== #
#  Sheet & structure tools
# ================================================================== #


@mcp.tool(
    annotations=ToolAnnotations(
        title="Manage Sheets",
        destructiveHint=True,
    ),
)
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


@mcp.tool(
    annotations=ToolAnnotations(
        title="Insert or Delete Rows/Columns",
        destructiveHint=True,
    ),
)
def insert_delete_cells(
    workbook: str,
    action: str,
    target: str = "row",
    position: int = 1,
    count: int = 1,
    sheet: str | None = None,
) -> dict:
    """Insert or delete rows/columns in a worksheet.

    Args:
        workbook: Workbook name or full file path.
        action: 'insert' or 'delete'.
        target: 'row' or 'column'.
        position: Row number or column number (1-based) where the action starts.
        count: Number of rows/columns to insert or delete.
        sheet: Sheet name. Uses the active sheet when omitted.
    """
    return _call(
        _handler.insert_delete_cells,
        workbook,
        action,
        sheet,
        position,
        count,
        target,
    )


# ================================================================== #
#  Formatting
# ================================================================== #


@mcp.tool(
    annotations=ToolAnnotations(
        title="Format Range",
        destructiveHint=True,
    ),
)
def format_range(
    workbook: str,
    cell_range: str,
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
        workbook: Workbook name or full file path.
        cell_range: Range like 'A1:D10'.
        sheet: Sheet name. Uses the active sheet when omitted.
        bold: Set bold.
        italic: Set italic.
        underline: Set underline.
        font_size: Font size in points.
        font_color: Font colour as hex string like '#FF0000'.
        bg_color: Background colour as hex string like '#FFFF00'.
        number_format: Excel number format like '#,##0.00'.
        alignment: Horizontal alignment: 'left', 'center', 'right', 'justify'.
        wrap_text: Enable text wrapping.
        border: Apply thin borders to all edges.
    """
    return _call(
        _handler.format_range,
        workbook,
        sheet,
        cell_range,
        bold,
        italic,
        underline,
        font_size,
        font_color,
        bg_color,
        number_format,
        alignment,
        wrap_text,
        border,
    )
