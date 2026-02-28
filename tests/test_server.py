"""Tests for MCP server tool registration.

NOTE: Full integration tests require Windows + Excel.
These tests verify the server structure only.
"""

from __future__ import annotations


def test_import_server():
    """Ensure the server module can be imported."""
    from mcp_server_xlwings.server import mcp  # noqa: F401


def test_server_has_tools():
    """Verify all 17 tools are registered."""
    from mcp_server_xlwings.server import mcp

    tool_names = {name for name in mcp._tool_manager._tools}
    expected = {
        # Active Excel (xlwings-exclusive)
        "get_active_workbook",
        "read_selection",
        "activate_sheet",
        "run_macro",
        "close_workbook",
        "recalculate",
        # Core workbook
        "list_open_workbooks",
        "open_workbook",
        "save_workbook",
        # Read / Write
        "read_range",
        "write_range",
        "read_cell_info",
        "set_formula",
        "find_replace",
        # Sheet & structure
        "manage_sheets",
        "insert_delete_cells",
        # Formatting
        "format_range",
    }
    assert expected == tool_names, (
        f"Missing: {expected - tool_names}, Extra: {tool_names - expected}"
    )
