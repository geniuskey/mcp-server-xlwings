"""Tests for MCP server tool registration.

NOTE: Full integration tests require Windows + Excel.
These tests verify the server structure only.
"""

from __future__ import annotations

import inspect


def test_import_server():
    """Ensure the server module can be imported."""
    from mcp_server_xlwings.server import mcp  # noqa: F401


def test_server_has_11_tools():
    """Verify all 11 tools are registered."""
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
        "get_formulas",
        "get_cell_styles",
        "get_objects",
    }
    assert expected == tool_names, (
        f"Missing: {expected - tool_names}, Extra: {tool_names - expected}"
    )


def test_read_data_optional_workbook():
    """read_data should have workbook as optional parameter."""
    from mcp_server_xlwings.server import read_data

    param = inspect.signature(read_data).parameters["workbook"]
    assert param.default is None


def test_read_data_merge_info_param():
    """read_data should have merge_info as optional bool (default False)."""
    from mcp_server_xlwings.server import read_data

    param = inspect.signature(read_data).parameters["merge_info"]
    assert param.default is False


def test_read_data_header_row_param():
    """read_data should have header_row as optional int (default None)."""
    from mcp_server_xlwings.server import read_data

    param = inspect.signature(read_data).parameters["header_row"]
    assert param.default is None


def test_write_data_exclusive_params():
    """write_data should have both data and formula as optional."""
    from mcp_server_xlwings.server import write_data

    sig = inspect.signature(write_data)
    assert sig.parameters["data"].default is None
    assert sig.parameters["formula"].default is None


def test_get_formulas_params():
    """get_formulas should have cell_range as required, values_too default False."""
    from mcp_server_xlwings.server import get_formulas

    sig = inspect.signature(get_formulas)
    assert sig.parameters["cell_range"].default is inspect.Parameter.empty
    assert sig.parameters["values_too"].default is False


def test_get_cell_styles_params():
    """get_cell_styles should have cell_range as required, properties default None."""
    from mcp_server_xlwings.server import get_cell_styles

    sig = inspect.signature(get_cell_styles)
    assert sig.parameters["cell_range"].default is inspect.Parameter.empty
    assert sig.parameters["properties"].default is None


def test_get_objects_params():
    """get_objects should have all parameters as optional."""
    from mcp_server_xlwings.server import get_objects

    sig = inspect.signature(get_objects)
    assert sig.parameters["workbook"].default is None
    assert sig.parameters["sheet"].default is None
