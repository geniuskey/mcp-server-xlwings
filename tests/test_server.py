"""Tests for MCP server tool registration.

NOTE: Full integration tests require Windows + Excel.
These tests verify the server structure only.
"""

from __future__ import annotations

import inspect


def test_import_server():
    """Ensure the server module can be imported."""
    from mcp_server_xlwings.server import mcp  # noqa: F401


def test_server_has_8_tools():
    """Verify all 8 consolidated tools are registered."""
    from mcp_server_xlwings.server import mcp

    tool_names = {name for name in mcp._tool_manager._tools}
    expected = {
        "get_active_workbook",
        "manage_workbooks",
        "read_data",
        "write_data",
        "manage_sheets",
        "find_replace",
        "format_range",
        "run_macro",
    }
    assert expected == tool_names, (
        f"Missing: {expected - tool_names}, Extra: {tool_names - expected}"
    )


def test_read_data_optional_workbook():
    """read_data should have workbook as optional parameter."""
    from mcp_server_xlwings.server import read_data

    param = inspect.signature(read_data).parameters["workbook"]
    assert param.default is None


def test_write_data_exclusive_params():
    """write_data should have both data and formula as optional."""
    from mcp_server_xlwings.server import write_data

    sig = inspect.signature(write_data)
    assert sig.parameters["data"].default is None
    assert sig.parameters["formula"].default is None
