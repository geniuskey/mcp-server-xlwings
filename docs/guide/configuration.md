# Configuration

Add `mcp-server-xlwings` to your MCP client. All clients use a similar JSON structure.

## Claude Desktop

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

## Claude Code

```bash
claude mcp add xlwings -- uvx mcp-server-xlwings
```

## Roo Code (VS Code)

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

## Cursor

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

## Windsurf

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

## Continue (VS Code)

Add to `~/.continue/config.yaml`:

```yaml
mcpServers:
  - name: xlwings
    command: uvx
    args:
      - mcp-server-xlwings
```

## Local Development

To use a local checkout instead of the PyPI package:

```json
{
  "mcpServers": {
    "xlwings": {
      "command": "uv",
      "args": [
        "run",
        "--directory", "/path/to/mcp-server-xlwings",
        "mcp-server-xlwings"
      ]
    }
  }
}
```
