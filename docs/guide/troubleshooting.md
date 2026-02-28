# Troubleshooting

## Common Errors

### "No active workbook found"

Excel is running but no workbook is open, or Excel is not running at all.

**Fix:** Open a workbook in Excel first, or use:
```json
{ "action": "open", "filepath": "C:\\path\\to\\file.xlsx" }
```

### "Workbook 'xxx' is not open and the path does not exist"

The specified workbook name doesn't match any open workbook, and the path doesn't exist on disk.

**Fix:** Check open workbooks with `manage_workbooks(action="list")`, or provide the full absolute path.

### "Sheet 'xxx' not found"

The sheet name doesn't match. Sheet names are **case-sensitive**.

**Fix:** Use `manage_sheets(action="list")` to see exact sheet names.

### "Excel COM error"

Excel process crashed, is unresponsive, or a dialog box is blocking.

**Fix:**
1. Check if Excel has a dialog box open (save prompt, error alert, etc.)
2. Close any dialogs, then retry
3. If Excel is frozen, restart it

## MCP Connection Issues

### Claude Desktop: "Disconnected"

1. Verify `claude_desktop_config.json` is valid JSON (no trailing commas)
2. Ensure `uvx` is in your system PATH
3. Restart Claude Desktop completely after config changes
4. Check logs: `%APPDATA%\Claude\logs\`

### Cursor / Windsurf: Server not starting

1. Verify the config file path is correct
2. Ensure Python 3.10+ is installed
3. Test manually: run `uvx mcp-server-xlwings` in a terminal
4. Check for error output in the terminal

### Claude Code: Tool not found

1. Verify with `claude mcp list`
2. Re-add if missing: `claude mcp add xlwings -- uvx mcp-server-xlwings`

## Limitations

| Limitation | Details |
|------------|---------|
| **Windows only** | COM automation requires Windows OS |
| **Excel required** | Microsoft Excel must be installed and licensed |
| **No headless mode** | Excel must be running (visible or minimized) — cannot run as a background service |
| **Single instance** | Connects to the active Excel application instance |
| **Concurrent access** | Multiple MCP clients connecting simultaneously may cause conflicts |
| **File size** | Very large files (100MB+) may be slow due to Excel's own memory usage |

## FAQ

### Can I use this on macOS or Linux?

No. Excel COM automation is Windows-only. For cross-platform needs, consider openpyxl-based MCP servers.

### Does it work with Excel Online or Google Sheets?

No. It requires the desktop Excel application running locally on Windows.

### Can I open password-protected files?

Yes, if Excel can open them. For files with password prompts, open them manually in Excel first, then use the MCP tools.

### Does it work with .xls (legacy format)?

Yes. Any format Excel can open works: `.xls`, `.xlsx`, `.xlsm`, `.xlsb`, `.csv`, `.tsv`.

### Can I use it while working in Excel?

Yes. The MCP server and the user share the same live Excel instance. You can continue editing while the agent reads or writes data.

### How do I update to the latest version?

```bash
uvx mcp-server-xlwings
```

`uvx` automatically fetches the latest version from PyPI each time. To verify:

```bash
uvx mcp-server-xlwings --version
```
