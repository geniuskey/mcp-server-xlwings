# Comparison

## mcp-server-xlwings vs File-Based Libraries

| Feature | mcp-server-xlwings | openpyxl | pandas |
|---------|:------------------:|:--------:|:------:|
| DRM-protected files | ✅ | ❌ | ❌ |
| Live Excel session | ✅ | ❌ | ❌ |
| Run VBA macros | ✅ | ❌ | ❌ |
| Read current selection | ✅ | ❌ | ❌ |
| Real-time formula results | ✅ | ❌ | ❌ |
| Recalculate workbook | ✅ | ❌ | ❌ |
| Merged cell detection | ✅ | ✅ | ❌ |
| Cell formatting (read) | ✅ | ✅ | ❌ |
| Cell formatting (write) | ✅ | ✅ | ❌ |
| Chart detection | ✅ | ✅ | ❌ |
| Cross-platform | ❌ | ✅ | ✅ |
| No Excel required | ❌ | ✅ | ✅ |
| Headless / CI usage | ❌ | ✅ | ✅ |

## When to Use What

### Use mcp-server-xlwings when:

- Files are **DRM-protected** or encrypted by enterprise security
- You need to interact with a **live Excel session**
- **VBA macros** need to be executed
- **Real-time formula** results are needed
- Users are **working in Excel simultaneously**
- You need to see the user's **current selection**

### Use openpyxl / pandas when:

- Running on **Linux or macOS**
- Processing files in **batch without Excel**
- **Server-side** processing without GUI
- **CI/CD pipelines** need to read spreadsheets
- Excel is **not installed** on the machine

## vs Other Excel MCP Servers

Most Excel MCP servers use openpyxl or similar file-based libraries. They read `.xlsx` files directly from disk, which means they **cannot**:

- Open DRM-protected or encrypted files
- Interact with the running Excel process
- Execute VBA macros or get live formula results
- Read the user's current selection
- Access files locked by another process

mcp-server-xlwings uses **COM automation** for full Excel integration — the same technology that powers VBA and Excel add-ins.
