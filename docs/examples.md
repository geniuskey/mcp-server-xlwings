# Examples

Real-world prompts showing how an AI agent uses each tool, with full request parameters and response JSON.

---

## Getting Started

### Explore a Workbook

> "What's in the spreadsheet I have open?"

**Step 1** -- The agent discovers the workbook structure:

```json
// Tool: get_active_workbook()

// Response
{
  "name": "sales_report_2024.xlsx",
  "path": "C:\\Users\\user\\Documents\\sales_report_2024.xlsx",
  "sheets": [
    { "name": "Revenue", "used_range": "$A$1:$N$45", "rows": 45, "columns": 14 },
    { "name": "Expenses", "used_range": "$A$1:$H$30", "rows": 30, "columns": 8 },
    { "name": "Dashboard", "used_range": "$A$1:$P$25", "rows": 25, "columns": 16 }
  ],
  "active_sheet": "Revenue",
  "selection": {
    "address": "$B$5:$D$10",
    "sheet": "Revenue",
    "data": [["Product", "Q1", "Q2"], ["Widget A", 12500, 13800], ["Widget B", 8400, 9100], ["Widget C", 15200, 16700], ["Widget D", 6300, 7200], ["Widget E", 11000, 12400]],
    "rows": 6,
    "columns": 3
  }
}
```

The agent now knows there are 3 sheets and can see the currently selected data. It decides to get more detail on the active sheet.

**Step 2** -- The agent reads the sheet summary to understand structure:

```json
// Tool: read_data()
// Parameters: (none -- omitting cell_range returns a summary)

// Response
{
  "sheet": "Revenue",
  "used_range": "$A$1:$N$45",
  "total_rows": 45,
  "total_columns": 14,
  "headers": ["Region", "Product", "Q1 Sales", "Q2 Sales", "Q3 Sales", "Q4 Sales", "Total", null, "YoY Growth"],
  "merged_cells": [
    { "range": "$A$1:$N$1", "value": "Annual Sales Report 2024", "rows": 1, "columns": 14 },
    { "range": "$A$3:$A$15", "value": "North", "rows": 13, "columns": 1 },
    { "range": "$A$16:$A$28", "value": "South", "rows": 13, "columns": 1 }
  ],
  "regions": [
    { "range": "$A$1:$N$45", "rows": 45, "columns": 14 }
  ]
}
```

The agent sees merged headers spanning multiple columns, merged region labels down column A, and that the data starts at row 3 with headers in row 2.

**Step 3** -- The agent reads the actual data:

```json
// Tool: read_data()
// Parameters: { "cell_range": "A2:G20", "merge_info": true }

// Response
{
  "range": "$A$2:$G$20",
  "sheet": "Revenue",
  "rows": 19,
  "columns": 7,
  "headers": ["Region", "Product", "Q1 Sales", "Q2 Sales", "Q3 Sales", "Q4 Sales", "Total"],
  "data": [
    ["North", "Widget A", 12500, 13800, 14200, 15600, 56100],
    ["North", "Widget B", 8400, 9100, 8900, 9800, 36200],
    ["North", "Widget C", 15200, 16700, 17100, 18300, 67300]
  ],
  "merged_ranges": [
    { "range": "$A$3:$A$15", "value": "North" }
  ]
}
```

With `merge_info=true`, every row under the merged "North" cell gets the value filled in, so the agent can process the data without null gaps.

---

### Summarize Selected Data

> "Summarize the data I've selected in Excel"

```json
// Tool: get_active_workbook()

// Response
{
  "name": "budget_2025.xlsx",
  "path": "C:\\Users\\user\\budget_2025.xlsx",
  "sheets": [
    { "name": "Operating", "used_range": "$A$1:$F$50", "rows": 50, "columns": 6 }
  ],
  "active_sheet": "Operating",
  "selection": {
    "address": "$A$1:$F$12",
    "sheet": "Operating",
    "data": [
      ["Category", "Jan", "Feb", "Mar", "Q1 Total", "Budget"],
      ["Payroll", 85000, 85000, 87000, 257000, 260000],
      ["Rent", 12000, 12000, 12000, 36000, 36000],
      ["Utilities", 3200, 3500, 2900, 9600, 10000],
      ["Software", 8500, 8500, 9200, 26200, 25000],
      ["Marketing", 15000, 22000, 18000, 55000, 50000],
      ["Travel", 4500, 6200, 3800, 14500, 15000],
      ["Supplies", 1200, 1400, 1100, 3700, 4000],
      ["Insurance", 5000, 5000, 5000, 15000, 15000],
      ["Training", 2000, 0, 3500, 5500, 8000],
      ["Misc", 800, 1200, 600, 2600, 3000],
      ["Total", 137200, 144800, 143100, 425100, 426000]
    ],
    "rows": 12,
    "columns": 6
  }
}
```

The selection data is returned directly by `get_active_workbook()` -- no additional call needed. The agent can immediately summarize: Q1 operating expenses totaled $425,100 against a $426,000 budget, with Marketing ($55,000) exceeding its $50,000 budget while most other categories came in under.

---

### Analyze Sheet Structure

> "What tables are in this sheet? It looks complicated."

```json
// Tool: read_data()
// Parameters: (none)

// Response
{
  "sheet": "Consolidated",
  "used_range": "$B$2:$T$85",
  "total_rows": 84,
  "total_columns": 19,
  "headers": ["Division", "Account", null, null, "2023 Actual", "2024 Budget", "2024 Actual", "Variance"],
  "merged_cells": [
    { "range": "$B$2:$T$2", "value": "FY2024 Consolidated Financial Summary", "rows": 1, "columns": 19 },
    { "range": "$B$4:$H$4", "value": "Income Statement", "rows": 1, "columns": 7 },
    { "range": "$B$40:$H$40", "value": "Balance Sheet", "rows": 1, "columns": 7 },
    { "range": "$J$4:$T$4", "value": "Cash Flow Statement", "rows": 1, "columns": 11 }
  ],
  "regions": [
    { "range": "$B$4:$H$38", "rows": 35, "columns": 7 },
    { "range": "$B$40:$H$85", "rows": 46, "columns": 7 },
    { "range": "$J$4:$T$60", "rows": 57, "columns": 11 }
  ]
}
```

The agent identifies three distinct tables: an Income Statement (B4:H38), a Balance Sheet (B40:H85), and a Cash Flow Statement (J4:T60). The merged cells at row 2 and row 4 serve as the report title and table headers respectively.

---

## Reading and Analysis

### Batch Read All Sheets

> "Give me a summary of every sheet in this workbook"

```json
// Tool: read_data()
// Parameters: { "sheet": "*" }

// Response
{
  "sheet_count": 4,
  "sheets": {
    "Revenue": {
      "sheet": "Revenue",
      "used_range": "$A$1:$N$45",
      "total_rows": 45,
      "total_columns": 14,
      "headers": ["Region", "Product", "Q1 Sales", "Q2 Sales", "Q3 Sales", "Q4 Sales", "Total"]
    },
    "Expenses": {
      "sheet": "Expenses",
      "used_range": "$A$1:$H$30",
      "total_rows": 30,
      "total_columns": 8,
      "headers": ["Department", "Category", "Amount", "Approved", "Paid", "Remaining"]
    },
    "Headcount": {
      "sheet": "Headcount",
      "used_range": "$A$1:$E$120",
      "total_rows": 120,
      "total_columns": 5,
      "headers": ["Employee ID", "Name", "Department", "Title", "Start Date"]
    },
    "Dashboard": {
      "sheet": "Dashboard",
      "used_range": "$A$1:$P$25",
      "total_rows": 25,
      "total_columns": 16,
      "headers": ["KPI", "Target", "Actual", "Status"]
    }
  }
}
```

A single call with `sheet="*"` returns structure summaries for all 4 sheets. The agent can now decide which sheets need deeper reading.

---

### Read Data with Merged Cells

> "Read the project schedule table -- the phase names are merged across multiple rows"

```json
// Tool: read_data()
// Parameters: { "cell_range": "B5:F18", "merge_info": true }

// Response
{
  "range": "$B$5:$F$18",
  "sheet": "Schedule",
  "rows": 14,
  "columns": 5,
  "headers": ["Phase", "Task", "Owner", "Start", "End"],
  "data": [
    ["Planning", "Requirements gathering", "Kim", "2024-01-08", "2024-01-19"],
    ["Planning", "Stakeholder review", "Park", "2024-01-22", "2024-01-26"],
    ["Planning", "Scope finalization", "Kim", "2024-01-29", "2024-02-02"],
    ["Development", "Backend API", "Lee", "2024-02-05", "2024-03-15"],
    ["Development", "Frontend UI", "Choi", "2024-02-12", "2024-03-22"],
    ["Development", "Integration", "Lee", "2024-03-25", "2024-04-05"],
    ["Testing", "QA testing", "Jung", "2024-04-08", "2024-04-26"],
    ["Testing", "UAT", "Park", "2024-04-29", "2024-05-10"],
    ["Deployment", "Staging release", "Lee", "2024-05-13", "2024-05-17"],
    ["Deployment", "Production release", "Lee", "2024-05-20", "2024-05-24"]
  ],
  "merged_ranges": [
    { "range": "$B$6:$B$8", "value": "Planning" },
    { "range": "$B$9:$B$11", "value": "Development" },
    { "range": "$B$12:$B$13", "value": "Testing" },
    { "range": "$B$14:$B$15", "value": "Deployment" }
  ]
}
```

Without `merge_info=true`, the Phase column would show "Planning" only in the first merged row and `null` for the rest. With it enabled, every row gets the correct phase label.

---

### Handle Non-Standard Header Positions

> "This sheet has a title block at the top. The actual column headers are on row 4."

```json
// Tool: read_data()
// Parameters: { "cell_range": "A1:H20", "header_row": 4 }

// Response
{
  "range": "$A$1:$H$20",
  "sheet": "Inventory",
  "rows": 20,
  "columns": 8,
  "headers": ["SKU", "Product Name", "Category", "Warehouse", "Qty On Hand", "Reorder Point", "Unit Cost", "Total Value"],
  "data": [
    ["SKU-001", "Industrial Valve 2in", "Plumbing", "Warehouse A", 340, 100, 45.50, 15470.00],
    ["SKU-002", "Copper Pipe 1in x 10ft", "Plumbing", "Warehouse A", 1200, 500, 12.75, 15300.00],
    ["SKU-003", "Circuit Breaker 20A", "Electrical", "Warehouse B", 85, 50, 28.90, 2456.50],
    ["SKU-004", "LED Panel Light 2x4", "Electrical", "Warehouse B", 210, 75, 62.00, 13020.00],
    ["SKU-005", "HVAC Filter 20x25", "HVAC", "Warehouse C", 450, 200, 8.50, 3825.00]
  ]
}
```

When a spreadsheet has a company logo, title, or metadata in rows 1-3, setting `header_row=4` tells the tool to treat row 4 as the column header row. Data is returned starting from row 5 onward with proper header labels.

---

### Find All Formulas

> "Show me all formulas in the Revenue sheet so I can audit the calculations"

```json
// Tool: get_formulas()
// Parameters: { "cell_range": "A1:N45", "sheet": "Revenue", "values_too": true }

// Response
{
  "formulas": [
    { "cell": "G3", "formula": "=SUM(C3:F3)", "value": 56100 },
    { "cell": "G4", "formula": "=SUM(C4:F4)", "value": 36200 },
    { "cell": "G5", "formula": "=SUM(C5:F5)", "value": 67300 },
    { "cell": "G15", "formula": "=SUM(G3:G14)", "value": 382400 },
    { "cell": "G28", "formula": "=SUM(G16:G27)", "value": 298700 },
    { "cell": "G30", "formula": "=G15+G28", "value": 681100 },
    { "cell": "H3", "formula": "=(G3-N3)/N3", "value": 0.0842 },
    { "cell": "H4", "formula": "=(G4-N4)/N4", "value": -0.0156 },
    { "cell": "N3", "formula": "='2023 Data'!G3", "value": 51740 }
  ],
  "total_formula_cells": 9,
  "range": "$A$1:$N$45",
  "sheet": "Revenue"
}
```

The agent finds SUM formulas for row totals (column G), growth-rate calculations referencing prior-year data (column H), and cross-sheet references pulling from the "2023 Data" sheet (column N). This gives a complete picture of the calculation logic.

---

### Identify Formatting Patterns with Styles

> "I need to understand the visual structure of this sheet. Which cells are headers vs data vs totals?"

```json
// Tool: get_cell_styles()
// Parameters: { "cell_range": "A1:H30", "sheet": "PnL", "properties": ["bold", "bg_color", "font_color", "font_size"] }

// Response
{
  "range": "$A$1:$H$30",
  "sheet": "PnL",
  "styles": [
    { "cell": "A1", "bold": true, "bg_color": "#1B3A5C", "font_color": "#FFFFFF", "font_size": 14 },
    { "cell": "B1", "bold": true, "bg_color": "#1B3A5C", "font_color": "#FFFFFF", "font_size": 14 },
    { "cell": "C1", "bold": true, "bg_color": "#1B3A5C", "font_color": "#FFFFFF", "font_size": 14 },
    { "cell": "A2", "bold": true, "bg_color": "#D6E4F0" },
    { "cell": "B2", "bold": true, "bg_color": "#D6E4F0" },
    { "cell": "A10", "bold": true, "bg_color": "#E2EFDA" },
    { "cell": "A15", "bold": true, "font_color": "#C00000" },
    { "cell": "A20", "bold": true, "bg_color": "#FFF2CC" },
    { "cell": "A25", "bold": true, "bg_color": "#1B3A5C", "font_color": "#FFFFFF" }
  ]
}
```

The agent can now deduce the visual hierarchy: dark blue background (#1B3A5C) with white text = top-level headers, light blue (#D6E4F0) = section headers, green (#E2EFDA) = subtotal rows, red font (#C00000) = negative/loss indicators, and yellow (#FFF2CC) = highlight rows. Only cells with non-default styling are returned, keeping the response compact.

---

### Combining Styles and Data to Understand Structure

> "Read the P&L sheet and help me figure out which rows are section headers vs data rows"

**Step 1** -- Read styles to identify header rows:

```json
// Tool: get_cell_styles()
// Parameters: { "cell_range": "A1:A30", "sheet": "PnL", "properties": ["bold", "bg_color"] }

// Response
{
  "range": "$A$1:$A$30",
  "sheet": "PnL",
  "styles": [
    { "cell": "A1", "bold": true, "bg_color": "#1B3A5C" },
    { "cell": "A2", "bold": true, "bg_color": "#D6E4F0" },
    { "cell": "A8", "bold": true, "bg_color": "#D6E4F0" },
    { "cell": "A14", "bold": true },
    { "cell": "A15", "bold": true, "bg_color": "#D6E4F0" },
    { "cell": "A22", "bold": true },
    { "cell": "A25", "bold": true, "bg_color": "#E2EFDA" }
  ]
}
```

**Step 2** -- Read the actual data in column A to see the labels:

```json
// Tool: read_data()
// Parameters: { "cell_range": "A1:H30", "sheet": "PnL" }

// Response
{
  "range": "$A$1:$H$30",
  "sheet": "PnL",
  "rows": 30,
  "columns": 8,
  "headers": ["Account", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "H1 Total"],
  "data": [
    ["Revenue", null, null, null, null, null, null, null],
    ["Product Sales", 125000, 132000, 128000, 141000, 138000, 145000, 809000],
    ["Service Revenue", 45000, 47000, 46000, 48000, 49000, 51000, 286000],
    ["License Fees", 18000, 18000, 18000, 18000, 18000, 18000, 108000],
    ["Other Income", 3200, 2800, 4100, 3600, 2900, 3800, 20400],
    ["Total Revenue", 191200, 199800, 196100, 210600, 207900, 217800, 1223400],
    ["Cost of Goods Sold", null, null, null, null, null, null, null],
    ["Direct Materials", 42000, 44500, 43200, 47800, 46500, 49000, 273000],
    ["Direct Labor", 31000, 31000, 31500, 32000, 32000, 33000, 190500],
    ["Manufacturing Overhead", 15000, 15200, 15100, 15800, 15600, 16000, 92700],
    ["Total COGS", 88000, 90700, 89800, 95600, 94100, 98000, 556200],
    ["Gross Profit", 103200, 109100, 106300, 115000, 113800, 119800, 667200],
    ["Operating Expenses", null, null, null, null, null, null, null]
  ]
}
```

By cross-referencing styles and data: rows with `bg_color=#D6E4F0` (A2, A8, A15) are section headers ("Revenue", "Cost of Goods Sold", "Operating Expenses"). Rows that are bold-only (A14, A22) are subtotals ("Total Revenue", "Total COGS"). Rows with green background (A25) are grand totals. Plain rows are data entries.

---

### Discover Dashboard Charts and Objects

> "What charts and images are on the Dashboard sheet?"

```json
// Tool: get_objects()
// Parameters: { "sheet": "Dashboard" }

// Response
{
  "sheet": "Dashboard",
  "charts": [
    {
      "name": "Chart 1",
      "top_left_cell": "$B$2",
      "width": 520,
      "height": 320,
      "chart_type": "5",
      "title": "Monthly Revenue Trend"
    },
    {
      "name": "Chart 2",
      "top_left_cell": "$J$2",
      "width": 400,
      "height": 320,
      "chart_type": "70",
      "title": "Revenue by Region"
    },
    {
      "name": "Chart 3",
      "top_left_cell": "$B$20",
      "width": 520,
      "height": 300,
      "chart_type": "51",
      "title": "Budget vs Actual"
    }
  ],
  "images": [
    { "name": "Picture 1", "top_left_cell": "$A$1", "width": 180, "height": 60 }
  ],
  "shapes": [
    { "name": "TextBox 1", "top_left_cell": "$J$20", "width": 400, "height": 40 }
  ]
}
```

The agent finds 3 charts (a line chart for revenue trend, a pie chart for regional breakdown, and a bar chart for budget comparison), a company logo image at A1, and a text box annotation. This helps the agent understand the dashboard layout and describe it to the user without needing to "see" the visuals.

---

## Writing and Formatting

### Write Data to a Range

> "Create a summary table in the Summary sheet starting at A1"

```json
// Tool: write_data()
// Parameters:
{
  "start_cell": "A1",
  "sheet": "Summary",
  "data": [
    ["Metric", "Q1", "Q2", "Q3", "Q4", "Annual"],
    ["Revenue", 580000, 625000, 610000, 695000, 2510000],
    ["COGS", 261000, 281000, 274000, 312000, 1128000],
    ["Gross Profit", 319000, 344000, 336000, 383000, 1382000],
    ["Operating Expenses", 195000, 202000, 198000, 215000, 810000],
    ["Net Income", 124000, 142000, 138000, 168000, 572000]
  ]
}

// Response
{
  "message": "Data written successfully to Summary",
  "start_cell": "A1",
  "written_range": "$A$1:$F$6",
  "rows": 6,
  "columns": 6
}
```

The 2D array is written starting from A1. The first sub-array becomes the header row.

---

### Write Formulas

> "Add SUM formulas to total each column, and a margin percentage formula"

```json
// Tool: write_data()
// Parameters: { "start_cell": "B7", "sheet": "Summary", "formula": "=SUM(B2:B6)" }

// Response
{
  "cell": "$B$7",
  "sheet": "Summary",
  "formula": "=SUM(B2:B6)",
  "calculated_value": 1479000
}
```

```json
// Tool: write_data()
// Parameters: { "start_cell": "B8", "sheet": "Summary", "formula": "=B6/B2" }

// Response
{
  "cell": "$B$8",
  "sheet": "Summary",
  "formula": "=B6/B2",
  "calculated_value": 0.2138
}
```

The calculated value is returned immediately so the agent can verify the formula is correct. The margin is 21.4% -- the agent can confirm this matches expectations.

---

### Format a Header Row

> "Make the header row bold, centered, with dark blue background and white text"

```json
// Tool: format_range()
// Parameters:
{
  "cell_range": "A1:F1",
  "sheet": "Summary",
  "bold": true,
  "alignment": "center",
  "bg_color": "#1B3A5C",
  "font_color": "#FFFFFF",
  "font_size": 11
}

// Response
{
  "range": "$A$1:$F$1",
  "sheet": "Summary",
  "applied": ["bold=True", "alignment=center", "bg_color=#1B3A5C", "font_color=#FFFFFF", "font_size=11"]
}
```

---

### Format Numbers and Apply Borders

> "Format the currency columns with comma separators and add borders to the whole table"

```json
// Tool: format_range()
// Parameters:
{
  "cell_range": "B2:F6",
  "sheet": "Summary",
  "number_format": "#,##0",
  "border": true
}

// Response
{
  "range": "$B$2:$F$6",
  "sheet": "Summary",
  "applied": ["number_format=#,##0", "border=True"]
}
```

```json
// Tool: format_range()
// Parameters:
{
  "cell_range": "B8:F8",
  "sheet": "Summary",
  "number_format": "0.0%",
  "italic": true
}

// Response
{
  "range": "$B$8:$F$8",
  "sheet": "Summary",
  "applied": ["number_format=0.0%", "italic=True"]
}
```

The data cells get comma-separated number formatting and thin borders. The margin row gets percentage formatting with italics to visually distinguish it.

---

### Insert Rows

> "Insert 3 blank rows above row 5 to add a new expense category"

```json
// Tool: manage_sheets()
// Parameters: { "action": "insert_rows", "sheet": "Summary", "position": 5, "count": 3 }

// Response
{
  "message": "Inserted 3 row(s) at row 5.",
  "sheet": "Summary"
}
```

Three blank rows are inserted at row 5, pushing existing rows 5+ downward. Formulas referencing those rows are automatically adjusted by Excel.

---

## Search and Management

### Find and Replace

> "Replace all instances of 'FY2024' with 'FY2025' in the report"

```json
// Tool: find_replace()
// Parameters: { "find": "FY2024", "replace": "FY2025" }

// Response
{
  "matches": [
    { "cell": "$A$1", "value": "FY2025 Annual Report", "sheet": "Cover" },
    { "cell": "$B$2", "value": "FY2025 Consolidated Financial Summary", "sheet": "Consolidated" },
    { "cell": "$A$1", "value": "FY2025 Revenue Detail", "sheet": "Revenue" },
    { "cell": "$D$3", "value": "FY2025 Budget", "sheet": "Budget" }
  ],
  "count": 4,
  "replaced": true
}
```

All 4 occurrences across multiple sheets were found and replaced. The response shows the updated values after replacement.

---

### Search Only (No Replace)

> "Find all cells that mention 'depreciation'"

```json
// Tool: find_replace()
// Parameters: { "find": "depreciation" }

// Response
{
  "matches": [
    { "cell": "$A$18", "value": "Depreciation & Amortization", "sheet": "PnL" },
    { "cell": "$B$5", "value": "Accumulated Depreciation", "sheet": "Balance Sheet" },
    { "cell": "$A$12", "value": "Depreciation", "sheet": "Cash Flow" }
  ],
  "count": 3
}
```

Omitting the `replace` parameter performs a search-only operation. The agent now knows exactly where depreciation appears.

---

### Manage Sheets

> "Create a new sheet called 'Q1 Report' and rename the old 'Sheet1' to 'Raw Data'"

```json
// Tool: manage_sheets()
// Parameters: { "action": "add", "new_name": "Q1 Report" }

// Response
{
  "message": "Sheet 'Q1 Report' added.",
  "sheet": "Q1 Report"
}
```

```json
// Tool: manage_sheets()
// Parameters: { "action": "rename", "sheet": "Sheet1", "new_name": "Raw Data" }

// Response
{
  "message": "Sheet 'Sheet1' renamed to 'Raw Data'.",
  "sheet": "Raw Data"
}
```

---

### Run a VBA Macro

> "Run the RefreshPivotTables macro to update all pivot tables"

```json
// Tool: run_macro()
// Parameters: { "macro_name": "RefreshPivotTables" }

// Response
{
  "message": "Macro 'RefreshPivotTables' executed successfully.",
  "return_value": "3 pivot tables refreshed"
}
```

The macro ran and returned a status message. If the macro accepts arguments, pass them via the `args` parameter as an array.

---

### Work with Multiple Workbooks

> "Copy the Q1 actuals from report.xlsx into the budget tracker"

**Step 1** -- Read data from the source workbook:

```json
// Tool: read_data()
// Parameters: { "workbook": "report.xlsx", "sheet": "Q1", "cell_range": "A1:E15" }

// Response
{
  "range": "$A$1:$E$15",
  "sheet": "Q1",
  "rows": 15,
  "columns": 5,
  "headers": ["Account", "Jan Actual", "Feb Actual", "Mar Actual", "Q1 Total"],
  "data": [
    ["Revenue", 191200, 199800, 196100, 587100],
    ["COGS", 88000, 90700, 89800, 268500],
    ["Gross Profit", 103200, 109100, 106300, 318600],
    ["SG&A", 62000, 65000, 63500, 190500],
    ["R&D", 28000, 29000, 28500, 85500],
    ["EBITDA", 13200, 15100, 14300, 42600]
  ]
}
```

**Step 2** -- Write to the destination workbook:

```json
// Tool: write_data()
// Parameters:
{
  "workbook": "budget_tracker.xlsx",
  "sheet": "Actuals",
  "start_cell": "A1",
  "data": [
    ["Account", "Jan Actual", "Feb Actual", "Mar Actual", "Q1 Total"],
    ["Revenue", 191200, 199800, 196100, 587100],
    ["COGS", 88000, 90700, 89800, 268500],
    ["Gross Profit", 103200, 109100, 106300, 318600],
    ["SG&A", 62000, 65000, 63500, 190500],
    ["R&D", 28000, 29000, 28500, 85500],
    ["EBITDA", 13200, 15100, 14300, 42600]
  ]
}

// Response
{
  "message": "Data written successfully to Actuals",
  "start_cell": "A1",
  "written_range": "$A$1:$E$7",
  "rows": 7,
  "columns": 5
}
```

The agent reads from one workbook and writes to another by specifying the `workbook` parameter. Both workbooks must be open in Excel.

---

## Real-World Workflows

### Workflow 1: Financial Statement Analysis

> "Analyze this business feasibility spreadsheet and extract the key financial metrics"

A complex Excel workbook with 5 sheets, merged headers, multi-level sections, and cross-sheet formulas. Here is how the agent processes it step-by-step.

**Step 1 -- Discover the workbook structure:**

```json
// Tool: get_active_workbook()

// Response
{
  "name": "feasibility_study_2024.xlsx",
  "path": "C:\\Users\\user\\Documents\\feasibility_study_2024.xlsx",
  "sheets": [
    { "name": "사업개요", "used_range": "$A$1:$L$35", "rows": 35, "columns": 12 },
    { "name": "매출계획", "used_range": "$A$1:$U$50", "rows": 50, "columns": 21 },
    { "name": "원가분석", "used_range": "$A$1:$N$40", "rows": 40, "columns": 14 },
    { "name": "손익계산", "used_range": "$A$1:$H$60", "rows": 60, "columns": 8 },
    { "name": "투자수익", "used_range": "$A$1:$F$25", "rows": 25, "columns": 6 }
  ],
  "active_sheet": "사업개요",
  "selection": {
    "address": "$A$1",
    "sheet": "사업개요",
    "data": [["사업 타당성 분석 보고서"]],
    "rows": 1,
    "columns": 1
  }
}
```

The agent sees 5 sheets: Business Overview (사업개요), Revenue Plan (매출계획), Cost Analysis (원가분석), P&L Statement (손익계산), and Investment Returns (투자수익). The Revenue Plan sheet is the largest at 50 rows x 21 columns.

**Step 2 -- Batch overview of all sheets:**

```json
// Tool: read_data()
// Parameters: { "sheet": "*" }

// Response
{
  "sheet_count": 5,
  "sheets": {
    "사업개요": {
      "sheet": "사업개요",
      "used_range": "$A$1:$L$35",
      "total_rows": 35,
      "total_columns": 12,
      "headers": ["항목", "내용", null, null, "비고"],
      "merged_cells": [
        { "range": "$A$1:$L$1", "value": "사업 타당성 분석 보고서", "rows": 1, "columns": 12 },
        { "range": "$A$3:$B$3", "value": "사업명", "rows": 1, "columns": 2 }
      ],
      "regions": [
        { "range": "$A$1:$L$35", "rows": 35, "columns": 12 }
      ]
    },
    "매출계획": {
      "sheet": "매출계획",
      "used_range": "$A$1:$U$50",
      "total_rows": 50,
      "total_columns": 21,
      "headers": ["구분", "단위", "면적(m2)", null, "분양가(만원)", null, "분양금액(백만원)"],
      "merged_cells": [
        { "range": "$A$1:$U$1", "value": "매출 계획표", "rows": 1, "columns": 21 },
        { "range": "$A$3:$A$4", "value": "구분", "rows": 2, "columns": 1 },
        { "range": "$A$6:$A$19", "value": "건축 계획", "rows": 14, "columns": 1 },
        { "range": "$A$20:$A$35", "value": "분양 수입", "rows": 16, "columns": 1 }
      ],
      "regions": [
        { "range": "$A$1:$U$50", "rows": 50, "columns": 21 }
      ]
    },
    "원가분석": {
      "sheet": "원가분석",
      "used_range": "$A$1:$N$40",
      "total_rows": 40,
      "total_columns": 14,
      "headers": ["비용항목", "산출근거", "금액(백만원)", "비율(%)"]
    },
    "손익계산": {
      "sheet": "손익계산",
      "used_range": "$A$1:$H$60",
      "total_rows": 60,
      "total_columns": 8,
      "headers": ["항목", "1차년도", "2차년도", "3차년도", "4차년도", "5차년도", "합계"]
    },
    "투자수익": {
      "sheet": "투자수익",
      "used_range": "$A$1:$F$25",
      "total_rows": 25,
      "total_columns": 6,
      "headers": ["지표", "값", "기준", "판정"]
    }
  }
}
```

The agent now has a map of the entire workbook. The Revenue Plan (매출계획) has multi-level merged headers with "건축 계획" (Building Plan) and "분양 수입" (Sales Revenue) spanning many rows. The P&L (손익계산) has a 5-year projection. The agent decides to examine the Revenue Plan headers first.

**Step 3 -- Understand the Revenue Plan header structure:**

```json
// Tool: read_data()
// Parameters: { "sheet": "매출계획", "cell_range": "A1:U5", "merge_info": true }

// Response
{
  "range": "$A$1:$U$5",
  "sheet": "매출계획",
  "rows": 5,
  "columns": 21,
  "headers": ["매출 계획표", "매출 계획표", "매출 계획표", "매출 계획표", "매출 계획표"],
  "data": [
    [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null],
    ["구분", "세부항목", "단위", "수량", "면적(m2)", "전용면적", "공용면적", "분양가(만원/m2)", "분양가(만원/평)", "분양금액(백만원)", "비율(%)", null, null, null, null, null, null, null, null, null, null],
    ["구분", "세부항목", "단위", "수량", "면적(m2)", "전용면적", "공용면적", "분양가(만원/m2)", "분양가(만원/평)", "분양금액(백만원)", "비율(%)", null, null, null, null, null, null, null, null, null, null]
  ],
  "merged_ranges": [
    { "range": "$A$1:$U$1", "value": "매출 계획표" },
    { "range": "$A$3:$A$4", "value": "구분" },
    { "range": "$B$3:$B$4", "value": "세부항목" }
  ]
}
```

The agent sees that rows 3-4 form a two-level header (구분/세부항목 span two rows). The actual data starts at row 5. Column layout: Category, Sub-item, Unit, Quantity, Area(m2), Exclusive Area, Common Area, Price per m2, Price per pyeong, Sales Amount, Ratio.

**Step 4 -- Find all formulas to understand calculation logic:**

```json
// Tool: get_formulas()
// Parameters: { "sheet": "매출계획", "cell_range": "A1:U50", "values_too": true }

// Response
{
  "formulas": [
    { "cell": "F6", "formula": "=D6*E6*0.75", "value": 5625.00 },
    { "cell": "G6", "formula": "=D6*E6*0.25", "value": 1875.00 },
    { "cell": "J6", "formula": "=D6*H6*E6/100", "value": 37828 },
    { "cell": "K6", "formula": "=J6/$J$36", "value": 0.3739 },
    { "cell": "J19", "formula": "=SUM(J6:J18)", "value": 101150 },
    { "cell": "J35", "formula": "=SUM(J20:J34)", "value": 28450 },
    { "cell": "J36", "formula": "=J19+J35", "value": 129600 },
    { "cell": "K36", "formula": "=J36/J36", "value": 1.0 }
  ],
  "total_formula_cells": 8,
  "range": "$A$1:$U$50",
  "sheet": "매출계획"
}
```

Key insights: Exclusive area is 75% of total (F6), common area is 25% (G6). Sales amounts are calculated as quantity x price x area (J6). Each line item's ratio is relative to the grand total at J36. Building plan subtotal is 101,150M won, other revenue is 28,450M won, for a grand total of 129,600M won (about 129.6 billion won).

**Step 5 -- Identify section structure via formatting:**

```json
// Tool: get_cell_styles()
// Parameters: { "sheet": "매출계획", "cell_range": "A1:A50", "properties": ["bold", "bg_color", "font_size"] }

// Response
{
  "range": "$A$1:$A$50",
  "sheet": "매출계획",
  "styles": [
    { "cell": "A1", "bold": true, "bg_color": "#1B3A5C", "font_size": 14 },
    { "cell": "A3", "bold": true, "bg_color": "#D6E4F0" },
    { "cell": "A6", "bold": true, "bg_color": "#F2F2F2" },
    { "cell": "A19", "bold": true, "bg_color": "#E2EFDA" },
    { "cell": "A20", "bold": true, "bg_color": "#F2F2F2" },
    { "cell": "A35", "bold": true, "bg_color": "#E2EFDA" },
    { "cell": "A36", "bold": true, "bg_color": "#1B3A5C" }
  ]
}
```

The formatting reveals the hierarchy: dark blue (#1B3A5C) = title and grand total rows, light blue (#D6E4F0) = column headers, light gray (#F2F2F2) = section start rows (건축 계획 at A6, 분양 수입 at A20), and green (#E2EFDA) = subtotal rows (A19, A35). Grand total at A36.

**Step 6 -- Extract the main data table:**

```json
// Tool: read_data()
// Parameters: { "sheet": "매출계획", "cell_range": "A5:K36", "header_row": 2, "merge_info": true }

// Response
{
  "range": "$A$5:$K$36",
  "sheet": "매출계획",
  "rows": 32,
  "columns": 11,
  "headers": ["구분", "세부항목", "단위", "수량", "면적(m2)", "전용면적", "공용면적", "분양가(만원/m2)", "분양가(만원/평)", "분양금액(백만원)", "비율(%)"],
  "data": [
    ["건축 계획", "아파트 59m2", "세대", 120, 59.0, 44.25, 14.75, 850, 2805, 37828, 0.3739],
    ["건축 계획", "아파트 84m2", "세대", 80, 84.0, 63.00, 21.00, 920, 3036, 30912, 0.2385],
    ["건축 계획", "아파트 115m2", "세대", 30, 115.0, 86.25, 28.75, 1050, 3465, 18113, 0.1397],
    ["건축 계획", "오피스텔 33m2", "실", 50, 33.0, 24.75, 8.25, 780, 2574, 6435, 0.0497],
    ["건축 계획", "상가 1층", "호", 15, 45.0, 33.75, 11.25, 1800, 5940, 6075, 0.0469],
    ["건축 계획", "상가 2층", "호", 10, 40.0, 30.00, 10.00, 950, 3135, 1900, 0.0147],
    ["건축 계획", "소계", null, null, null, null, null, null, null, 101150, 0.7806],
    ["분양 수입", "발코니 확장", "세대", 200, null, null, null, null, null, 10000, 0.0772],
    ["분양 수입", "시스템 에어컨", "세대", 150, null, null, null, null, null, 4500, 0.0347],
    ["분양 수입", "주차장 수입", "면", 350, null, null, null, null, null, 8750, 0.0675],
    ["분양 수입", "기타 수입", null, null, null, null, null, null, null, 5200, 0.0401],
    ["분양 수입", "소계", null, null, null, null, null, null, null, 28450, 0.2195],
    ["합계", null, null, null, null, null, null, null, null, 129600, 1.0]
  ]
}
```

Using `header_row=2` correctly maps row 3-4's merged headers. With `merge_info=true`, the "건축 계획" and "분양 수입" merged labels fill every row. The agent now has the complete revenue plan: total projected sales of 129,600M won (~129.6 billion won), with 78% from building sales and 22% from ancillary revenue.

**Step 7 -- Read the Investment Returns sheet for key metrics:**

```json
// Tool: read_data()
// Parameters: { "sheet": "투자수익", "cell_range": "A1:F25" }

// Response
{
  "range": "$A$1:$F$25",
  "sheet": "투자수익",
  "rows": 25,
  "columns": 6,
  "headers": ["지표", "값", "기준", "판정"],
  "data": [
    ["총 사업비", "118,500백만원", null, null],
    ["총 분양수입", "129,600백만원", null, null],
    ["사업이익", "11,100백만원", null, null],
    ["사업이익률", "9.37%", "8% 이상", "적합"],
    ["IRR (내부수익률)", "12.8%", "10% 이상", "적합"],
    ["NPV (순현재가치)", "8,520백만원", "0 이상", "적합"],
    ["투자회수기간", "3.2년", "5년 이내", "적합"],
    ["손익분기점", "78.5%", "85% 이하", "적합"],
    ["자기자본수익률(ROE)", "18.6%", "15% 이상", "적합"]
  ]
}
```

The agent can now provide a comprehensive summary: Total project cost is 118.5B won with projected revenue of 129.6B won, yielding an 11.1B won profit (9.37% margin). All investment metrics pass their thresholds -- IRR of 12.8%, NPV of 8.52B won, payback period of 3.2 years, break-even at 78.5% occupancy, and ROE of 18.6%.

---

### Workflow 2: Multi-Workbook Data Consolidation

> "I have Q1, Q2, and Q3 sales reports in separate workbooks. Consolidate them into a yearly summary workbook."

The agent needs to read from three source workbooks and write a consolidated summary into a fourth. All four workbooks must be open in Excel.

**Step 1 -- List open workbooks:**

```json
// Tool: manage_workbooks()
// Parameters: { "action": "list" }

// Response
[
  {
    "name": "sales_q1.xlsx",
    "path": "C:\\Users\\user\\Reports\\sales_q1.xlsx",
    "sheets": [{ "name": "Sales", "used_range": "$A$1:$F$25", "rows": 25, "columns": 6 }]
  },
  {
    "name": "sales_q2.xlsx",
    "path": "C:\\Users\\user\\Reports\\sales_q2.xlsx",
    "sheets": [{ "name": "Sales", "used_range": "$A$1:$F$25", "rows": 25, "columns": 6 }]
  },
  {
    "name": "sales_q3.xlsx",
    "path": "C:\\Users\\user\\Reports\\sales_q3.xlsx",
    "sheets": [{ "name": "Sales", "used_range": "$A$1:$F$25", "rows": 25, "columns": 6 }]
  },
  {
    "name": "annual_summary.xlsx",
    "path": "C:\\Users\\user\\Reports\\annual_summary.xlsx",
    "sheets": [{ "name": "Summary", "used_range": "$A$1", "rows": 1, "columns": 1 }]
  }
]
```

All four workbooks are open. Each quarterly report has the same structure (25 rows, 6 columns).

**Step 2 -- Read Q1 data:**

```json
// Tool: read_data()
// Parameters: { "workbook": "sales_q1.xlsx", "sheet": "Sales", "cell_range": "A1:F25" }

// Response
{
  "range": "$A$1:$F$25",
  "sheet": "Sales",
  "rows": 25,
  "columns": 6,
  "headers": ["Region", "Product", "Units Sold", "Revenue", "Cost", "Profit"],
  "data": [
    ["North", "Widget A", 1250, 312500, 187500, 125000],
    ["North", "Widget B", 840, 168000, 100800, 67200],
    ["North", "Widget C", 1520, 456000, 273600, 182400],
    ["South", "Widget A", 980, 245000, 147000, 98000],
    ["South", "Widget B", 1100, 220000, 132000, 88000],
    ["South", "Widget C", 760, 228000, 136800, 91200],
    ["East", "Widget A", 1350, 337500, 202500, 135000],
    ["East", "Widget B", 920, 184000, 110400, 73600],
    ["East", "Widget C", 1180, 354000, 212400, 141600],
    ["West", "Widget A", 1050, 262500, 157500, 105000],
    ["West", "Widget B", 680, 136000, 81600, 54400],
    ["West", "Widget C", 890, 267000, 160200, 106800]
  ]
}
```

**Step 3 -- Read Q2 and Q3 data:**

```json
// Tool: read_data()
// Parameters: { "workbook": "sales_q2.xlsx", "sheet": "Sales", "cell_range": "A1:F25" }

// Response
{
  "range": "$A$1:$F$25",
  "sheet": "Sales",
  "rows": 25,
  "columns": 6,
  "headers": ["Region", "Product", "Units Sold", "Revenue", "Cost", "Profit"],
  "data": [
    ["North", "Widget A", 1380, 345000, 207000, 138000],
    ["North", "Widget B", 910, 182000, 109200, 72800],
    ["North", "Widget C", 1670, 501000, 300600, 200400],
    ["South", "Widget A", 1050, 262500, 157500, 105000],
    ["South", "Widget B", 1200, 240000, 144000, 96000],
    ["South", "Widget C", 820, 246000, 147600, 98400],
    ["East", "Widget A", 1420, 355000, 213000, 142000],
    ["East", "Widget B", 980, 196000, 117600, 78400],
    ["East", "Widget C", 1250, 375000, 225000, 150000],
    ["West", "Widget A", 1120, 280000, 168000, 112000],
    ["West", "Widget B", 720, 144000, 86400, 57600],
    ["West", "Widget C", 950, 285000, 171000, 114000]
  ]
}
```

```json
// Tool: read_data()
// Parameters: { "workbook": "sales_q3.xlsx", "sheet": "Sales", "cell_range": "A1:F25" }

// Response
{
  "range": "$A$1:$F$25",
  "sheet": "Sales",
  "rows": 25,
  "columns": 6,
  "headers": ["Region", "Product", "Units Sold", "Revenue", "Cost", "Profit"],
  "data": [
    ["North", "Widget A", 1420, 355000, 213000, 142000],
    ["North", "Widget B", 890, 178000, 106800, 71200],
    ["North", "Widget C", 1710, 513000, 307800, 205200],
    ["South", "Widget A", 1080, 270000, 162000, 108000],
    ["South", "Widget B", 1150, 230000, 138000, 92000],
    ["South", "Widget C", 850, 255000, 153000, 102000],
    ["East", "Widget A", 1480, 370000, 222000, 148000],
    ["East", "Widget B", 1010, 202000, 121200, 80800],
    ["East", "Widget C", 1300, 390000, 234000, 156000],
    ["West", "Widget A", 1150, 287500, 172500, 115000],
    ["West", "Widget B", 750, 150000, 90000, 60000],
    ["West", "Widget C", 980, 294000, 176400, 117600]
  ]
}
```

The agent now has all three quarters of data. It aggregates by region and product, computing totals.

**Step 4 -- Write the consolidated summary with headers:**

```json
// Tool: write_data()
// Parameters:
{
  "workbook": "annual_summary.xlsx",
  "sheet": "Summary",
  "start_cell": "A1",
  "data": [
    ["Region", "Product", "Q1 Units", "Q2 Units", "Q3 Units", "YTD Units", "Q1 Revenue", "Q2 Revenue", "Q3 Revenue", "YTD Revenue", "YTD Profit", "Profit Margin"],
    ["North", "Widget A", 1250, 1380, 1420, 4050, 312500, 345000, 355000, 1012500, 405000, "40.0%"],
    ["North", "Widget B", 840, 910, 890, 2640, 168000, 182000, 178000, 528000, 211200, "40.0%"],
    ["North", "Widget C", 1520, 1670, 1710, 4900, 456000, 501000, 513000, 1470000, 588000, "40.0%"],
    ["South", "Widget A", 980, 1050, 1080, 3110, 245000, 262500, 270000, 777500, 311000, "40.0%"],
    ["South", "Widget B", 1100, 1200, 1150, 3450, 220000, 240000, 230000, 690000, 276000, "40.0%"],
    ["South", "Widget C", 760, 820, 850, 2430, 228000, 246000, 255000, 729000, 291600, "40.0%"],
    ["East", "Widget A", 1350, 1420, 1480, 4250, 337500, 355000, 370000, 1062500, 425000, "40.0%"],
    ["East", "Widget B", 920, 980, 1010, 2910, 184000, 196000, 202000, 582000, 232800, "40.0%"],
    ["East", "Widget C", 1180, 1250, 1300, 3730, 354000, 375000, 390000, 1119000, 447600, "40.0%"],
    ["West", "Widget A", 1050, 1120, 1150, 3320, 262500, 280000, 287500, 830000, 332000, "40.0%"],
    ["West", "Widget B", 680, 720, 750, 2150, 136000, 144000, 150000, 430000, 172000, "40.0%"],
    ["West", "Widget C", 890, 950, 980, 2820, 267000, 285000, 294000, 846000, 338400, "40.0%"]
  ]
}

// Response
{
  "message": "Data written successfully to Summary",
  "start_cell": "A1",
  "written_range": "$A$1:$L$13",
  "rows": 13,
  "columns": 12
}
```

**Step 5 -- Add a totals row with formulas:**

```json
// Tool: write_data()
// Parameters:
{
  "workbook": "annual_summary.xlsx",
  "sheet": "Summary",
  "start_cell": "A14",
  "data": [["Grand Total"]]
}

// Response
{
  "message": "Data written successfully to Summary",
  "start_cell": "A14",
  "written_range": "$A$14:$A$14",
  "rows": 1,
  "columns": 1
}
```

```json
// Tool: write_data()
// Parameters: { "workbook": "annual_summary.xlsx", "sheet": "Summary", "start_cell": "F14", "formula": "=SUM(F2:F13)" }

// Response
{
  "cell": "$F$14",
  "sheet": "Summary",
  "formula": "=SUM(F2:F13)",
  "calculated_value": 39760
}
```

```json
// Tool: write_data()
// Parameters: { "workbook": "annual_summary.xlsx", "sheet": "Summary", "start_cell": "J14", "formula": "=SUM(J2:J13)" }

// Response
{
  "cell": "$J$14",
  "sheet": "Summary",
  "formula": "=SUM(J2:J13)",
  "calculated_value": 10076500
}
```

```json
// Tool: write_data()
// Parameters: { "workbook": "annual_summary.xlsx", "sheet": "Summary", "start_cell": "K14", "formula": "=SUM(K2:K13)" }

// Response
{
  "cell": "$K$14",
  "sheet": "Summary",
  "formula": "=SUM(K2:K13)",
  "calculated_value": 4030600
}
```

**Step 6 -- Format the summary table:**

```json
// Tool: format_range()
// Parameters:
{
  "workbook": "annual_summary.xlsx",
  "sheet": "Summary",
  "cell_range": "A1:L1",
  "bold": true,
  "alignment": "center",
  "bg_color": "#1B3A5C",
  "font_color": "#FFFFFF",
  "wrap_text": true
}

// Response
{
  "range": "$A$1:$L$1",
  "sheet": "Summary",
  "applied": ["bold=True", "alignment=center", "bg_color=#1B3A5C", "font_color=#FFFFFF", "wrap_text=True"]
}
```

```json
// Tool: format_range()
// Parameters:
{
  "workbook": "annual_summary.xlsx",
  "sheet": "Summary",
  "cell_range": "G2:K14",
  "number_format": "#,##0"
}

// Response
{
  "range": "$G$2:$K$14",
  "sheet": "Summary",
  "applied": ["number_format=#,##0"]
}
```

```json
// Tool: format_range()
// Parameters:
{
  "workbook": "annual_summary.xlsx",
  "sheet": "Summary",
  "cell_range": "A14:L14",
  "bold": true,
  "bg_color": "#E2EFDA",
  "border": true
}

// Response
{
  "range": "$A$14:$L$14",
  "sheet": "Summary",
  "applied": ["bold=True", "bg_color=#E2EFDA", "border=True"]
}
```

**Step 7 -- Save the result:**

```json
// Tool: manage_workbooks()
// Parameters: { "action": "save", "workbook": "annual_summary.xlsx" }

// Response
{
  "message": "Workbook 'annual_summary.xlsx' saved.",
  "path": "C:\\Users\\user\\Reports\\annual_summary.xlsx"
}
```

The consolidation is complete. Data from three quarterly workbooks has been aggregated into a formatted summary with 12 product-region combinations, YTD totals of 39,760 units and $10,076,500 in revenue, and a consistent 40% profit margin across all product lines. The summary includes SUM formulas so it stays correct if individual values are updated.
