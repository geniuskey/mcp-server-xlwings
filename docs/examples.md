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

## Read All Sheets at Once

> "Give me a summary of every sheet in this workbook"

Agent calls `read_data(sheet="*")` -- returns summaries of all sheets in a single call.

## Read Data with Merged Cells

> "Read B6:C20 and fill in the merged cell values"

Agent calls `read_data(cell_range="B6:C20", merge_info=true)`. Merged cells are filled with the parent cell's value instead of returning null.

## Find All Formulas

> "Show me all formulas in the revenue sheet"

Agent calls `get_formulas(cell_range="A1:Z100", sheet="Revenue", values_too=true)` -- returns every formula with its calculated value.

## Identify Headers by Style

> "Which cells are bold or have a colored background?"

Agent calls `get_cell_styles(cell_range="A1:Z10", properties=["bold", "bg_color"])` -- returns only cells with non-default bold or background color.

## Check for Charts and Images

> "Are there any charts or images in this sheet?"

Agent calls `get_objects()` -- lists all charts, images, and shapes with their positions and sizes.

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

## Real-World Workflow: Financial Analysis

> "Analyze this business feasibility spreadsheet and extract the key financial metrics"

A complex Excel file with 5 sheets, merged headers, and formulas. Here's how the agent processes it step-by-step:

**Step 1 — Discover structure:**

```
get_active_workbook()
```

Returns workbook name, 5 sheets with their dimensions and used ranges.

**Step 2 — Batch overview:**

```
read_data(sheet="*")
```

Gets summaries of all 5 sheets in one call: headers, merged cells, regions.

**Step 3 — Understand headers:**

```
read_data(sheet="Revenue", cell_range="A1:Z5", merge_info=true)
```

Reveals multi-level merged headers like "Building Plan" spanning rows 6-19.

**Step 4 — Find calculations:**

```
get_formulas(sheet="Revenue", cell_range="A1:Z100", values_too=true)
```

Identifies all formula cells with their calculated values.

**Step 5 — Identify sections by formatting:**

```
get_cell_styles(sheet="Revenue", cell_range="A1:A50", properties=["bold", "bg_color"])
```

Bold + colored background = section headers. Bold only = subtotals.

**Step 6 — Extract target data:**

```
read_data(sheet="Revenue", cell_range="B6:U36", header_row=2)
```

Reads the specific data table with row 2 as headers.
