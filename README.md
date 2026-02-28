# mcp-server-xlwings

MCP server for Excel automation via xlwings COM. Works with DRM-protected files.

[![PyPI](https://img.shields.io/pypi/v/mcp-server-xlwings.svg)](https://pypi.org/project/mcp-server-xlwings/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/)

## Why xlwings?

Libraries like `openpyxl` or `pandas` read `.xlsx` files directly from disk. This fails when:

- **DRM / file-level encryption** is applied (common in enterprise environments)
- You need to interact with a **live Excel session** (formulas, macros, add-ins)
- Files are locked by another process

`mcp-server-xlwings` uses **COM automation** to talk to the running Excel process, so it can read and write any file that Excel itself can open -- including DRM-protected documents.

### xlwings-exclusive capabilities

These features are **impossible** with file-based libraries like openpyxl:

- **Read the user's current selection** -- see exactly what the user is looking at
- **Get the active workbook** -- no need to specify a file path
- **Run VBA macros** -- execute existing macros and get their return values
- **Live formula results** -- set a formula and get the calculated value immediately
- **Force recalculation** -- trigger Excel to recalculate all formulas

## Installation

### With uvx (recommended)

```bash
uvx mcp-server-xlwings
```

### With pip

```bash
pip install mcp-server-xlwings
```

## Configuration

### Claude Desktop

Add to `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "xlwings": {
      "command": "uvx",
      "args": ["mcp-server-xlwings"]
    }
  }
}
```

### Claude Code

```bash
claude mcp add xlwings -- uvx mcp-server-xlwings
```

### Roo Code (VS Code)

Add to Roo Code MCP settings or create `<project-root>/.roo/mcp.json`:

```json
{
  "mcpServers": {
    "xlwings": {
      "command": "uvx",
      "args": ["mcp-server-xlwings"]
    }
  }
}
```

### Cursor

Add to `%USERPROFILE%\.cursor\mcp.json`:

```json
{
  "mcpServers": {
    "xlwings": {
      "command": "uvx",
      "args": ["mcp-server-xlwings"]
    }
  }
}
```

### Windsurf

Add to `%USERPROFILE%\.codeium\windsurf\mcp_config.json`:

```json
{
  "mcpServers": {
    "xlwings": {
      "command": "uvx",
      "args": ["mcp-server-xlwings"]
    }
  }
}
```

### Continue (VS Code)

Add to `~/.continue/config.yaml`:

```yaml
mcpServers:
  - name: xlwings
    command: uvx
    args:
      - mcp-server-xlwings
```

## Available Tools (8)

All tools default to the **active workbook** when `workbook` is omitted.

| Tool | Description |
|------|-------------|
| `get_active_workbook` | Get active workbook info, sheets, and current selection with data |
| `manage_workbooks` | List, open, save, close, or recalculate workbooks |
| `read_data` | Read a cell range. Set `detail=True` on a single cell for formula/format info |
| `write_data` | Write a 2D array (`data`) or a single-cell formula (`formula`) |
| `manage_sheets` | List, add, delete, rename, copy, activate sheets. Insert/delete rows and columns |
| `find_replace` | Search for text, optionally replace it |
| `format_range` | Apply formatting (bold, italic, color, borders, alignment, number format, etc.) |
| `run_macro` | Execute a VBA macro and get its return value |

## Examples

### See what the user is working on

> "What's in the spreadsheet I have open?"

The agent calls `get_active_workbook()` to get the workbook name, sheets, and selection data, then `read_data()` to fetch the full sheet.

### Summarize selected data

> "Summarize the data I've selected"

The agent calls `get_active_workbook()` -- the response includes the selection data directly.

### Run a macro

> "Run the UpdateReport macro"

The agent calls `run_macro(macro_name="UpdateReport")` and returns the result.

### Build a summary row

> "Add a SUM formula in C10 that totals C2:C9"

The agent calls `write_data(start_cell="C10", formula="=SUM(C2:C9)")` and gets back the calculated value.

### Format a header row

> "Make row 1 bold and centered with a yellow background"

The agent calls `format_range(cell_range="A1:D1", bold=true, alignment="center", bg_color="#FFFF00")`.

### Insert rows

> "Insert 3 blank rows at row 5"

The agent calls `manage_sheets(action="insert_rows", position=5, count=3)`.

## Requirements

- **Windows** (Excel COM automation is Windows-only)
- **Microsoft Excel** installed
- **Python 3.10+**

## License

MIT
