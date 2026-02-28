"""Tests for ExcelHandler.

NOTE: These tests require a running Excel instance on Windows.
They are skipped automatically in CI / non-Windows environments.
"""

from __future__ import annotations

import platform
import pytest

pytestmark = pytest.mark.skipif(
    platform.system() != "Windows",
    reason="Excel COM automation requires Windows",
)


def test_import():
    """Ensure the module can be imported."""
    from mcp_server_xlwings.excel import ExcelHandler  # noqa: F401
