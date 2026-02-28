# Performance Guide

## How It Works

Each tool call flows through: **MCP protocol → Python → COM → Excel → response**. Individual COM calls take 10–50ms, so the key to performance is minimizing COM round-trips per tool invocation.

## Best Practices

### 1. Start with Sheet Summary

Always call `read_data()` without `cell_range` first:

```json
{ "sheet": "Sheet1" }
```

This returns dimensions, headers, merged cells, and regions with only ~5 COM calls — regardless of sheet size. Use the summary to decide which ranges to read next.

### 2. Use Batch Operations

| Instead of... | Use... | Savings |
|---------------|--------|---------|
| 5× `read_data()` per sheet | 1× `read_data(sheet="*")` | 5 calls → 1 |
| N× `read_data(detail=true)` per cell | 1× `get_formulas(cell_range="A1:Z100")` | N calls → 1 |
| N× checking cell formatting | 1× `get_cell_styles(cell_range="...")` | N calls → 1 |

### 3. Filter Style Properties

`get_cell_styles` makes multiple COM calls per cell. Use `properties` to query only what you need:

```json
{ "cell_range": "A1:Z50", "properties": ["bold", "bg_color"] }
```

This skips font_name, font_size, alignment, border, etc. — significantly reducing COM calls.

### 4. Use merge_info Selectively

`merge_info=true` checks every cell in the range for merge status (~1 COM call per cell). For large ranges (1000+ cells), this can be slow. Use it on **specific sections**, not entire sheets.

### 5. Large File Strategy

For workbooks with thousands of rows:

1. `read_data()` → get sheet summary (dimensions, headers)
2. `read_data(cell_range="A1:Z20")` → read header area to understand structure
3. Decide on target ranges based on the structure
4. Read specific ranges only — avoid reading the entire sheet

## Performance Reference

| Operation | Typical Time | COM Calls |
|-----------|:------------:|:---------:|
| Sheet summary (`read_data()`) | ~50ms | ~5 |
| Read 100 rows (bulk) | ~30ms | 1 |
| Read 1,000 rows (bulk) | ~100ms | 1 |
| `merge_info` on 100 cells | ~200ms | ~100 |
| `get_formulas` on 1,000 cells | ~50ms | 2 |
| `get_cell_styles` on 100 cells | ~500ms | ~100 × N props |
| `get_objects` | ~20ms | ~3 |
