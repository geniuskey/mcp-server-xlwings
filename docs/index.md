---
layout: home

hero:
  name: mcp-server-xlwings
  text: Excel Automation for AI Agents
  tagline: MCP server that controls live Excel via COM. Works with DRM-protected files. Windows only.
  image:
    src: /logo.svg
    alt: mcp-server-xlwings
  actions:
    - theme: brand
      text: Get Started
      link: /guide/getting-started
    - theme: alt
      text: View on GitHub
      link: https://github.com/geniuskey/mcp-server-xlwings

features:
  - icon: 🔒
    title: DRM-Protected Files
    details: Uses COM automation to talk to the running Excel process. Opens any file that Excel can open, including encrypted and DRM-protected documents.
  - icon: ⚡
    title: Live Excel Control
    details: Read selections, run VBA macros, get live formula results, and force recalculation — features impossible with file-based libraries.
  - icon: 🛠️
    title: 11 Powerful Tools
    details: Consolidated tool set covering read, write, format, search, formulas, styles, charts, and macro execution.
  - icon: 📊
    title: Smart Sheet Analysis
    details: Automatic merged cell detection, region discovery, and sheet structure analysis for complex enterprise spreadsheets.
  - icon: 🔄
    title: Batch Operations
    details: Read all sheets at once with sheet="*", get formulas in bulk, and query cell styles across ranges in a single call.
  - icon: 🚀
    title: Zero Config
    details: One-line setup with uvx. Works with Claude Desktop, Claude Code, Cursor, Windsurf, Roo Code, and Continue.
---

## Quick Start

Add to your MCP client config and start using immediately:

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

Then ask your AI agent:

> "What's in the spreadsheet I have open?"

The agent will call `get_active_workbook()` and `read_data()` to explore your Excel file.

[Get Started →](/guide/getting-started)
