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

## Available Tools

| Tool | Description |
|------|-------------|
| `list_open_workbooks` | List all currently open workbooks |
| `open_workbook` | Open a file or create a new workbook |
| `read_range` | Read data from a cell range |
| `write_range` | Write a 2D array to a cell range |
| `read_cell_info` | Get detailed cell info (value, formula, format) |
| `set_formula` | Set a formula and get the calculated result |
| `manage_sheets` | List, add, delete, rename, or copy sheets |
| `find_replace` | Search for text, optionally replace it |
| `format_range` | Apply formatting (bold, color, borders, etc.) |
| `save_workbook` | Save or Save As |

## Examples

### Read a report

> "Read the data from Sheet1 of report.xlsx"

The agent calls `read_range(workbook="report.xlsx")` and receives a structured 2D array with headers.

### Build a summary row

> "Add a SUM formula in C10 that totals C2:C9"

The agent calls `set_formula(workbook="report.xlsx", cell="C10", formula="=SUM(C2:C9)")` and gets back the calculated value.

### Format a header row

> "Make row 1 bold with a yellow background"

The agent calls `format_range(workbook="report.xlsx", cell_range="A1:D1", bold=true, bg_color="#FFFF00")`.

## Requirements

- **Windows** (Excel COM automation is Windows-only)
- **Microsoft Excel** installed
- **Python 3.10+**

## License

MIT
