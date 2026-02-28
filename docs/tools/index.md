# Tools Reference

All tools default to the **active workbook** when `workbook` is omitted.

## get_active_workbook

Get the currently active workbook info including per-sheet metadata, and current selection with its data.

**Parameters:** None

**Response:**

```json
{
  "name": "report.xlsx",
  "path": "C:\\Users\\user\\report.xlsx",
  "sheets": [
    { "name": "Data", "used_range": "$A$1:$Z$100", "rows": 100, "columns": 26 },
    { "name": "Summary", "used_range": "$A$1:$D$10", "rows": 10, "columns": 4 }
  ],
  "active_sheet": "Data",
  "selection": {
    "address": "$A$1:$C$5",
    "sheet": "Data",
    "data": [["ID", "Name", "Value"], [1, "Alice", 100], ...],
    "rows": 5,
    "columns": 3
  }
}
```

---

## manage_workbooks

Manage Excel workbooks: list, open, save, close, or recalculate.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `action` | string | Yes | `list`, `open`, `save`, `close`, `recalculate` |
| `workbook` | string | No | Workbook name or path. Defaults to active workbook |
| `filepath` | string | No | For `open`: file path (use `"new"` for blank). For `save`: Save As path |
| `read_only` | bool | No | For `open`: open in read-only mode (default `true`) |
| `save` | bool | No | For `close`: save before closing |

**Example -- list open workbooks:**

```json
// Request
{ "action": "list" }

// Response
[
  {
    "name": "report.xlsx",
    "path": "C:\\Users\\user\\report.xlsx",
    "sheets": [
      { "name": "Sheet1", "used_range": "$A$1:$D$50", "rows": 50, "columns": 4 }
    ]
  }
]
```

---

## read_data

Read data from an Excel range. When `cell_range` is omitted, returns a **sheet summary** with structure analysis instead of reading all data.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `workbook` | string | No | Defaults to active workbook |
| `sheet` | string | No | Defaults to active sheet. Use `*` to read all sheets |
| `cell_range` | string | No | Range like `A1:D10`. Returns sheet summary if omitted |
| `headers` | bool | No | Treat first row as headers (default `true`) |
| `detail` | bool | No | For single cells: include formula, type, format info |
| `merge_info` | bool | No | Fill merged cells with the merge area's value instead of null (default `false`) |
| `header_row` | int | No | 1-based row number to use as headers (e.g. `3` means row 3) |

**Example -- sheet summary (no range):**

```json
// Request
{ }

// Response
{
  "sheet": "Sheet1",
  "used_range": "$B$3:$Z$100",
  "total_rows": 98,
  "total_columns": 25,
  "headers": ["2024 Revenue", null, null, "Q1", "Q2", "Q3"],
  "merged_cells": [
    { "range": "$B$3:$H$3", "value": "2024 Revenue", "rows": 1, "columns": 7 },
    { "range": "$B$4:$D$4", "value": "Q1", "rows": 1, "columns": 3 }
  ],
  "regions": [
    { "range": "$B$3:$H$20", "rows": 18, "columns": 7 },
    { "range": "$B$25:$F$40", "rows": 16, "columns": 5 }
  ]
}
```

**Example -- read a specific range:**

```json
// Request
{ "cell_range": "A1:D10" }

// Response
{
  "range": "$A$1:$D$10",
  "sheet": "Sheet1",
  "rows": 10,
  "columns": 4,
  "headers": ["ID", "Name", "Date", "Amount"],
  "data": [
    [1, "Alice", "2024-01-15", 1500],
    [2, "Bob", "2024-01-16", 2300]
  ]
}
```

**Example -- single cell detail:**

```json
// Request
{ "cell_range": "C10", "detail": true }

// Response
{
  "range": "$C$10",
  "sheet": "Sheet1",
  "rows": 1,
  "columns": 1,
  "data": [[3800]],
  "detail": {
    "value": 3800,
    "type": "number",
    "formula": "=SUM(C2:C9)",
    "number_format": "#,##0",
    "font": { "name": "Calibri", "size": 11, "bold": true, "italic": false }
  }
}
```

**Example -- merge_info:**

```json
// Request
{ "cell_range": "B6:C8", "merge_info": true }

// Response
{
  "range": "$B$6:$C$8",
  "sheet": "Sheet1",
  "rows": 3,
  "columns": 2,
  "headers": ["건축 계획", "대지"],
  "data": [
    ["건축 계획", "면적"],
    ["건축 계획", null]
  ],
  "merged_ranges": [
    { "range": "$B$6:$B$19", "value": "건축 계획" }
  ]
}
```

**Example -- batch read all sheets:**

```json
// Request
{ "sheet": "*" }

// Response
{
  "sheet_count": 3,
  "sheets": {
    "Sheet1": { "sheet": "Sheet1", "used_range": "$A$1:$D$50", "total_rows": 50, "total_columns": 4, "headers": ["ID", "Name", "Date", "Amount"] },
    "Sheet2": { "sheet": "Sheet2", "used_range": "$A$1:$B$20", "total_rows": 20, "total_columns": 2, "headers": ["Category", "Total"] },
    "Summary": { "sheet": "Summary", "used_range": "$A$1:$C$5", "total_rows": 5, "total_columns": 3, "headers": ["Metric", "Value", "Change"] }
  }
}
```

---

## write_data

Write data or a formula to Excel cells. Provide `data` for a 2D array, or `formula` for a single-cell formula.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `start_cell` | string | Yes | Top-left cell (e.g. `A1`) |
| `data` | array | No | 2D list of values. Mutually exclusive with `formula` |
| `formula` | string | No | Excel formula like `=SUM(A1:A10)`. Mutually exclusive with `data` |
| `workbook` | string | No | Defaults to active workbook |
| `sheet` | string | No | Defaults to active sheet |

**Example -- write data:**

```json
// Request
{
  "start_cell": "A1",
  "data": [["Name", "Score"], ["Alice", 95], ["Bob", 87]]
}

// Response
{
  "message": "Data written successfully to Sheet1",
  "start_cell": "A1",
  "written_range": "$A$1:$B$3",
  "rows": 3,
  "columns": 2
}
```

**Example -- set formula:**

```json
// Request
{ "start_cell": "C10", "formula": "=SUM(C2:C9)" }

// Response
{
  "cell": "$C$10",
  "sheet": "Sheet1",
  "formula": "=SUM(C2:C9)",
  "calculated_value": 3800
}
```

---

## manage_sheets

Manage sheets and structure: list, add, delete, rename, copy, activate, insert/delete rows and columns.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `action` | string | Yes | `list`, `add`, `delete`, `rename`, `copy`, `activate`, `insert_rows`, `delete_rows`, `insert_columns`, `delete_columns` |
| `workbook` | string | No | Defaults to active workbook |
| `sheet` | string | No | Target sheet (required for delete/rename/copy/activate) |
| `new_name` | string | No | New name (for rename; optional for add/copy) |
| `position` | int | No | Row/column number (1-based) for insert/delete (default `1`) |
| `count` | int | No | Number of rows/columns (default `1`) |

**Example -- insert rows:**

```json
// Request
{ "action": "insert_rows", "position": 5, "count": 3 }

// Response
{ "message": "Inserted 3 row(s) at row 5.", "sheet": "Sheet1" }
```

---

## find_replace

Search for text in a sheet, optionally replacing it.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `find` | string | Yes | Text to search for |
| `workbook` | string | No | Defaults to active workbook |
| `sheet` | string | No | Defaults to active sheet |
| `replace` | string | No | Replacement text. Omit for search only |
| `match_case` | bool | No | Case-sensitive matching (default `false`) |

**Example -- search:**

```json
// Request
{ "find": "Alice" }

// Response
{
  "matches": [
    { "cell": "$A$2", "value": "Alice", "sheet": "Sheet1" },
    { "cell": "$A$15", "value": "Alice Smith", "sheet": "Sheet1" }
  ],
  "count": 2
}
```

---

## format_range

Apply formatting to a cell range.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `cell_range` | string | Yes | Range like `A1:D10` |
| `workbook` | string | No | Defaults to active workbook |
| `sheet` | string | No | Defaults to active sheet |
| `bold` | bool | No | Set bold |
| `italic` | bool | No | Set italic |
| `underline` | bool | No | Set underline |
| `font_size` | int | No | Font size in points |
| `font_color` | string | No | Hex colour like `#FF0000` |
| `bg_color` | string | No | Background hex colour like `#FFFF00` |
| `number_format` | string | No | Excel format like `#,##0.00` |
| `alignment` | string | No | `left`, `center`, `right`, `justify` |
| `wrap_text` | bool | No | Enable text wrapping |
| `border` | bool | No | Apply thin borders |

**Example:**

```json
// Request
{ "cell_range": "A1:D1", "bold": true, "alignment": "center", "bg_color": "#FFFF00" }

// Response
{
  "range": "$A$1:$D$1",
  "sheet": "Sheet1",
  "applied": ["bold=True", "alignment=center", "bg_color=#FFFF00"]
}
```

---

## run_macro

Execute a VBA macro in Excel and return its result.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `macro_name` | string | Yes | Macro name (e.g. `MyMacro` or `Module1.MyMacro`) |
| `workbook` | string | No | Workbook name. If omitted, Excel resolves globally |
| `args` | array | No | Arguments to pass to the macro |

**Example:**

```json
// Request
{ "macro_name": "UpdateReport" }

// Response
{
  "message": "Macro 'UpdateReport' executed successfully.",
  "return_value": "Report updated"
}
```

---

## get_formulas

Get all formulas in a range. Returns only cells that contain formulas.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `cell_range` | string | Yes | Range like `A1:U99` |
| `workbook` | string | No | Defaults to active workbook |
| `sheet` | string | No | Defaults to active sheet |
| `values_too` | bool | No | Include calculated values alongside formulas (default `false`) |

**Example:**

```json
// Request
{ "cell_range": "A1:U99", "values_too": true }

// Response
{
  "formulas": [
    { "cell": "J22", "formula": "=E22*H22", "value": 37828200 },
    { "cell": "K22", "formula": "=J22/$J$36", "value": 0.3739 }
  ],
  "total_formula_cells": 2,
  "range": "$A$1:$U$99",
  "sheet": "Sheet1"
}
```

---

## get_cell_styles

Get formatting/style info for cells in a range. Returns only cells with non-default styles. Useful for identifying headers, subtotals, and data roles by visual formatting.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `cell_range` | string | Yes | Range like `A1:D10` |
| `workbook` | string | No | Defaults to active workbook |
| `sheet` | string | No | Defaults to active sheet |
| `properties` | array | No | Filter specific properties: `bold`, `italic`, `underline`, `font_name`, `font_size`, `font_color`, `bg_color`, `number_format`, `alignment`, `border` |

**Example:**

```json
// Request
{ "cell_range": "A1:D5", "properties": ["bold", "bg_color", "font_color"] }

// Response
{
  "range": "$A$1:$D$5",
  "sheet": "Sheet1",
  "styles": [
    { "bold": true, "bg_color": "#1B3A5C", "font_color": "#FFFFFF", "cell": "A1" },
    { "bold": true, "bg_color": "#1B3A5C", "font_color": "#FFFFFF", "cell": "B1" },
    { "bold": true, "cell": "A5" }
  ]
}
```

---

## get_objects

List charts, images, and shapes on a sheet.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `workbook` | string | No | Defaults to active workbook |
| `sheet` | string | No | Defaults to active sheet |

**Example:**

```json
// Request
{ }

// Response
{
  "sheet": "Dashboard",
  "charts": [
    {
      "name": "Chart 1",
      "top_left_cell": "$F$2",
      "width": 480,
      "height": 300,
      "chart_type": "5",
      "title": "Monthly Revenue"
    }
  ],
  "images": [
    { "name": "Picture 1", "top_left_cell": "$A$1", "width": 200, "height": 80 }
  ],
  "shapes": []
}
```
