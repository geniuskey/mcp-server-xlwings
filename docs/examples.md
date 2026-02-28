# Examples

Real-world prompts and how the agent uses the tools.

## Explore a Spreadsheet

> "What's in the spreadsheet I have open?"

1. Agent calls `get_active_workbook()` -- gets workbook name, per-sheet metadata (rows, columns, used range), and current selection data
2. Agent calls `read_data()` -- gets sheet summary with headers, merged cells, and regions
3. Agent calls `read_data(cell_range="A1:D20")` -- fetches the specific data range

## Summarize Selected Data

> "Summarize the data I've selected"

Agent calls `get_active_workbook()` -- the response includes the selection data directly. No additional calls needed.

## Analyze a Complex Sheet

> "What tables are in this sheet?"

Agent calls `read_data()` without a range. The response includes:
- `used_range`: where data exists
- `merged_cells`: hierarchical header structure
- `regions`: distinct table boundaries

```json
{
  "used_range": "$B$3:$Z$100",
  "merged_cells": [
    { "range": "$B$3:$H$3", "value": "2024 Revenue", "rows": 1, "columns": 7 }
  ],
  "regions": [
    { "range": "$B$3:$H$20", "rows": 18, "columns": 7 },
    { "range": "$B$25:$F$40", "rows": 16, "columns": 5 }
  ]
}
```

## Build a Summary Row

> "Add a SUM formula in C10 that totals C2:C9"

Agent calls `write_data(start_cell="C10", formula="=SUM(C2:C9)")` and gets back the calculated value immediately.

## Format a Header Row

> "Make row 1 bold and centered with a yellow background"

Agent calls `format_range(cell_range="A1:D1", bold=true, alignment="center", bg_color="#FFFF00")`.

## Insert Rows

> "Insert 3 blank rows at row 5"

Agent calls `manage_sheets(action="insert_rows", position=5, count=3)`.

## Run a VBA Macro

> "Run the UpdateReport macro"

Agent calls `run_macro(macro_name="UpdateReport")` and returns the result.

## Work with Multiple Workbooks

> "Copy data from report.xlsx to summary.xlsx"

1. Agent calls `read_data(workbook="report.xlsx", cell_range="A1:D50")`
2. Agent calls `write_data(workbook="summary.xlsx", start_cell="A1", data=...)`

## Find and Replace

> "Replace all occurrences of '2024' with '2025'"

Agent calls `find_replace(find="2024", replace="2025")`.

## Manage Sheets

> "Create a new sheet called 'Q1 Report'"

Agent calls `manage_sheets(action="add", new_name="Q1 Report")`.
