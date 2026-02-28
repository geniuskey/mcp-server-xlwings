# Getting Started

## Requirements

- **Windows** (Excel COM automation is Windows-only)
- **Microsoft Excel** installed
- **Python 3.10+**

## Installation

### With uvx (recommended)

No installation needed. The MCP client runs it automatically:

```bash
uvx mcp-server-xlwings
```

### With pip

```bash
pip install mcp-server-xlwings
```

## Quick Start

1. Install the server (or use `uvx`)
2. Add the configuration to your MCP client (see [Configuration](./configuration))
3. Open an Excel file
4. Ask your AI agent to read or modify the spreadsheet

The server connects to the running Excel process via COM automation. **Excel must be running** for the tools to work.

## How It Works

```
AI Agent  <-->  MCP Client  <-->  mcp-server-xlwings  <-->  Excel (COM)
```

Unlike file-based libraries (openpyxl, pandas), this server talks to the **live Excel process**. This means:

- DRM-protected files work seamlessly
- Formulas return calculated values
- VBA macros can be executed
- The user's current selection is accessible
- Changes appear in real-time in the Excel window
