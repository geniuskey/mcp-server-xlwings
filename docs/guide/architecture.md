# Architecture

## Overview

```
┌─────────────────┐    stdio     ┌────────────────────┐     COM      ┌─────────┐
│   MCP Client    │ ───────────→ │ mcp-server-xlwings │ ───────────→ │  Excel   │
│ (Claude, Cursor │ ←─────────── │    (Python/MCP)    │ ←─────────── │ Process  │
│  Windsurf, etc.)│   JSON-RPC   └────────────────────┘  Win32 API   └─────────┘
└─────────────────┘
```

## Layers

### MCP Client

Claude Desktop, Claude Code, Cursor, Windsurf, Roo Code, Continue — any client that supports the Model Context Protocol. Sends tool calls via JSON-RPC over stdio.

### MCP Server

A Python process running FastMCP. Receives tool calls, translates them into xlwings operations, and returns structured JSON responses. Fully stateless — no session management needed.

### COM Automation (xlwings)

Connects to the running Excel process via Windows COM (the same interface VBA macros use). Can access any file Excel has open, including DRM-protected and encrypted documents.

### Excel Process

The actual Microsoft Excel application. Manages files, formulas, macros, and formatting. The user can continue working in Excel while the MCP server operates — they share the same live instance.

## Design Principles

### Stateless

No session IDs or connection management. Each tool call independently finds the active Excel instance. If Excel is restarted between calls, the next call simply reconnects.

### Active Workbook Default

All tools default to the active workbook when `workbook` is omitted. This eliminates the need to specify file paths for the most common case — working with the file the user currently has open.

### Summary-First

`read_data()` without a `cell_range` returns sheet structure (dimensions, headers, merged cells, regions) instead of dumping all data. This lets the agent understand the sheet layout before fetching specific ranges, saving context window space.

### Minimal Tool Count

11 tools cover the full range of Excel operations. Related actions are consolidated (e.g., `manage_workbooks` handles list/open/save/close/recalculate) to reduce LLM context overhead.
