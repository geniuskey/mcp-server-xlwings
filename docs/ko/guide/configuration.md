# 설정

MCP 클라이언트에 `mcp-server-xlwings`를 추가합니다. 모든 클라이언트가 유사한 JSON 구조를 사용합니다.

## Claude Desktop

`%APPDATA%\Claude\claude_desktop_config.json`에 추가합니다:

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

Roo Code MCP 설정에 추가하거나 `<project-root>/.roo/mcp.json`을 생성합니다:

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

`%USERPROFILE%\.cursor\mcp.json`에 추가합니다:

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

`%USERPROFILE%\.codeium\windsurf\mcp_config.json`에 추가합니다:

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

`~/.continue/config.yaml`에 추가합니다:

```yaml
mcpServers:
  - name: xlwings
    command: uvx
    args:
      - mcp-server-xlwings
```

## 로컬 개발

PyPI 패키지 대신 로컬 체크아웃을 사용하려면:

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
