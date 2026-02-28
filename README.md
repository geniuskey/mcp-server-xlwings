# mcp-server-xlwings

MCP server for Excel automation via xlwings COM. Works with DRM-protected files.

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
- **Live formula results** -- `set_formula` returns the calculated value immediately
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

Add to your `claude_desktop_config.json`:

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

Or with a local Python installation:

```json
{
  "mcpServers": {
    "xlwings": {
      "command": "python",
      "args": ["-m", "mcp_server_xlwings"]
    }
  }
}
```

### Claude Code

```bash
claude mcp add xlwings -- uvx mcp-server-xlwings
```

## Available Tools (17)

### Active Excel (xlwings-exclusive)

| Tool | Description |
|------|-------------|
| `get_active_workbook` | Get info about the currently active workbook, sheet, and selection |
| `read_selection` | Read data from the cells the user currently has selected |
| `activate_sheet` | Switch the visible sheet in Excel |
| `run_macro` | Execute a VBA macro and get its return value |
| `close_workbook` | Close a workbook, optionally saving first |
| `recalculate` | Force recalculation of all formulas |

### Core Workbook

| Tool | Description |
|------|-------------|
| `list_open_workbooks` | List all currently open workbooks |
| `open_workbook` | Open a file or create a new workbook |
| `save_workbook` | Save or Save As |

### Read / Write

| Tool | Description |
|------|-------------|
| `read_range` | Read data from a cell range |
| `write_range` | Write a 2D array to a cell range |
| `read_cell_info` | Get detailed cell info (value, formula, format, type) |
| `set_formula` | Set a formula and get the calculated result immediately |
| `find_replace` | Search for text, optionally replace it |

### Sheet & Structure

| Tool | Description |
|------|-------------|
| `manage_sheets` | List, add, delete, rename, or copy sheets |
| `insert_delete_cells` | Insert or delete rows/columns |

### Formatting

| Tool | Description |
|------|-------------|
| `format_range` | Apply formatting (bold, italic, underline, color, borders, alignment, wrap text, number format) |

## Examples

### See what the user is working on

> "What's in the spreadsheet I have open?"

The agent calls `get_active_workbook()` to discover the workbook name and sheets, then `read_range()` to fetch the data.

### Read the user's selection

> "Summarize the data I've selected"

The agent calls `read_selection()` to get exactly the cells the user has highlighted.

### Run a macro

> "Run the UpdateReport macro"

The agent calls `run_macro(macro_name="UpdateReport")` and returns the result.

### Build a summary row

> "Add a SUM formula in C10 that totals C2:C9"

The agent calls `set_formula(workbook="report.xlsx", cell="C10", formula="=SUM(C2:C9)")` and gets back the calculated value.

### Format a header row

> "Make row 1 bold and centered with a yellow background"

The agent calls `format_range(workbook="report.xlsx", cell_range="A1:D1", bold=true, alignment="center", bg_color="#FFFF00")`.

### Insert rows

> "Insert 3 blank rows at row 5"

The agent calls `insert_delete_cells(workbook="report.xlsx", action="insert", target="row", position=5, count=3)`.

## Requirements

- **Windows** (Excel COM automation is Windows-only)
- **Microsoft Excel** installed
- **Python 3.10+**

## License

MIT
