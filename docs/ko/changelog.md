# 변경 이력

## 0.4.0 (2026-02-28)

- `read_data` enhanced: `merge_info=true` fills merged cells with values + returns `merged_ranges`
- `read_data` enhanced: `header_row` parameter to specify which row is the header (1-based)
- `read_data` enhanced: `sheet="*"` batch-reads all sheets in one call
- New tool: `get_formulas` -- batch query all formulas in a range with optional calculated values
- New tool: `get_cell_styles` -- batch query cell formatting (bold, colors, number format, etc.)
- New tool: `get_objects` -- list charts, images, and shapes on a sheet
- Tool count: 8 → 11

## 0.3.0 (2026-02-28)

- `read_data()` without `cell_range` returns sheet summary instead of full data dump
- Merged cell detection in header area (first 20 rows) with range, value, and span info
- Distinct table region detection via CurrentRegion probing
- `get_active_workbook` and `manage_workbooks(action="list")` include per-sheet metadata (used range, rows, columns)
- README updated with setup guides for 6 MCP clients (Claude Desktop, Claude Code, Roo Code, Cursor, Windsurf, Continue)

## 0.2.0 (2026-02-28)

- **Breaking**: Consolidated 17 tools into 8 for faster Claude Desktop response
- `workbook` parameter now optional across all tools (defaults to active workbook)
- `get_active_workbook` includes selection data (merged `read_selection`)
- `manage_workbooks`: list, open, save, close, recalculate
- `read_data`: merged `read_range` + `read_cell_info` (use `detail=True`)
- `write_data`: merged `write_range` + `set_formula` (use `data` or `formula`)
- `manage_sheets`: expanded with activate, insert/delete rows/columns
- Fix `FastMCP` compatibility with mcp>=1.26.0 (`description` → `instructions`)
- Added PyPI trusted publishing workflow

## 0.1.0 (2026-02-28)

- Initial release
- 10 MCP tools for Excel automation via xlwings COM
- Sessionless design -- no session ID management
- Supports DRM-protected files through live Excel process control
