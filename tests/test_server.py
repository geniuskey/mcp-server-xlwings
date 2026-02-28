"""Tests for MCP server tool registration.

NOTE: Full integration tests require Windows + Excel.
These tests verify the server structure only.
"""

from __future__ import annotations


def test_import_server():
    """Ensure the server module can be imported."""
    from mcp_server_xlwings.server import mcp  # noqa: F401


def test_server_has_tools():
    """Verify all 10 tools are registered."""
    from mcp_server_xlwings.server import mcp

    tool_names = {name for name in mcp._tool_manager._tools}
    expected = {
        "list_open_workbooks",
        "open_workbook",
        "read_range",
        "write_range",
        "read_cell_info",
        "set_formula",
        "manage_sheets",
        "find_replace",
        "format_range",
        "save_workbook",
    }
    assert expected == tool_names, f"Missing tools: {expected - tool_names}"
