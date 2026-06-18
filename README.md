# plexcel — CSV to XLSX Workflow

```
  ██████  ██      ███████ ██   ██  ██████ ███████ ██
  ██   ██ ██      ██       ██ ██  ██      ██      ██
  ██████  ██      █████     ███   ██      █████   ██
  ██      ██      ██       ██ ██  ██      ██      ██
  ██      ███████ ███████ ██   ██  ██████ ███████ ███████
                  pLaInTeXt → eXcEl
```

Convert CSV files into styled Excel workbooks. Keep your data in plain-text CSVs, define styling in YAML, and generate `.xlsx` files on demand.

## Quick Start

```bash
# 1. Create venv and install dependencies (inside this repo)
python3 -m venv venv
./venv/bin/pip install -r requirements.txt

# 2. Add to your .zshrc (adjust path to where you cloned this repo)
export PLEXCEL_DIR="/path/to/sheets-styling"
source "$PLEXCEL_DIR/plexcel-functions.zsh"

# 3. Convert CSVs to Excel
csv-to-xlsx output.xlsx workbook.yaml
```

Or run the Python script directly:
```bash
./venv/bin/python csv_to_xlsx.py output.xlsx workbook.yaml
```

## How It Works

```
┌─────────────┐     ┌──────────────┐     ┌────────────┐
│  CSV files  │ ──► │ workbook.yaml│ ──► │ output.xlsx│
│ (your data) │     │ (layout/style│     │ (styled!)  │
└─────────────┘     └──────────────┘     └────────────┘
                           │
                    ┌──────┴───────┐
                    │ styles.yaml  │
                    │ (colors/fonts│
                    └──────────────┘
```

You maintain three things:
1. **CSV files** — your raw data (edit in any text editor or spreadsheet)
2. **workbook.yaml** — which CSVs become sheets, and how they're configured
3. **styles.yaml** — reusable named styles (colors, fonts, borders)

## Usage

### Config-based (recommended)
```bash
csv-to-xlsx output.xlsx workbook.yaml
```

### Simple one-liner (no config file needed)
```bash
csv-to-xlsx output.xlsx sales.csv:Sales team.csv:Team
```

### Utility commands
```bash
plexcel-init my-config.yaml   # Create a starter workbook.yaml
plexcel-styles                 # Edit your global styles.yaml
```

## workbook.yaml Reference

This is where you define your sheets. Each entry maps a CSV to a sheet with optional styling:

```yaml
sheets:
  - csv: sales.csv              # Path to CSV file
    name: Sales Report          # Sheet tab name in Excel
    freeze_pane: A2             # Lock rows/columns for scrolling
    auto_filter: true           # Add filter dropdowns to headers
    column_widths: {A: 25, B: 15, C: 12}
    row_heights: {"1": 30}
    wrap_columns: [E, F]        # These columns get text wrapping
    merged_cells: [A1:D1]       # Merge cell ranges
    cell_styles:                # Apply named styles to ranges
      A1:D1: header
    dropdowns:                  # Data validation lists
      - range: B2:B100
        values: North,South,East,West
    formulas:                   # Computed cells
      - column: D
        start_row: 2
        formula: "=B{row}*C{row}"
    number_formats:             # Excel number formatting
      - column: C
        start_row: 2
        format: "$#,##0.00"
    conditional_formatting:     # Color cells by value
      - type: cell_is
        range: D2:D100
        operator: greaterThan
        formula: ["10000"]
        style: {fill_color: FF92D050}
```

### All sheet options

| Option | Example | Description |
|--------|---------|-------------|
| `csv` | `data.csv` | CSV file path (required) |
| `name` | `My Sheet` | Sheet tab name |
| `freeze_pane` | `A2` | Freeze above/left of this cell |
| `freeze_row` | `1` | Freeze first N rows (alternative) |
| `auto_filter` | `true` / `row 3` / `A1:D1` | Enable filter dropdowns |
| `column_widths` | `{A: 30, B: 20}` | Set column widths |
| `row_heights` | `{"1": 30, "2": 25}` | Set specific row heights |
| `wrap_columns` | `[E, F]` | Enable text wrap on columns |
| `merged_cells` | `[A1:D1]` | Merge cell ranges |
| `cell_styles` | `{A1:D1: header}` | Apply named styles |
| `row_styles` | see below | Style rows matching a condition |
| `dropdowns` | see below | Data validation dropdowns |
| `formulas` | see below | Computed cell values |
| `number_formats` | see below | Excel number formats |
| `conditional_formatting` | see below | Color cells conditionally |
| `highlight` | see below | Simplified conditional coloring |

## styles.yaml Reference

Define reusable styles that you reference by name in `workbook.yaml`:

```yaml
defaults:
  font: Arial 10
  row_height: 25.5
  date_format: yyyy-mm-dd
  align: left center
  freeze_row: 1
  auto_filter: true

styles:
  header:
    font: Arial 14 bold white
    fill: "#1F3864"
    align: center center
    border: {bottom: thin}

  subheader:
    font: Arial 10 bold white
    fill: "#2E5597"
    align: center center
```

### Shorthand syntax

| Property | Shorthand | Example |
|----------|-----------|---------|
| `font` | `[family] [size] [bold] [italic] [color]` | `Arial 12 bold white` |
| `fill` | hex color or name | `"#1F3864"` or `red` |
| `align` | `[horizontal] [vertical] [wrap]` | `center center wrap` |
| `border` | style or `{edge: style}` | `thin` or `{bottom: thin}` |

### Named colors

`white`, `black`, `red`, `green`, `blue`, `yellow`, `orange`, `purple`, `gray`

### Border styles

`thin`, `medium`, `thick`, `double`, `dotted`, `dashed`, `hair`

## Feature Details

### Freeze Panes

```yaml
freeze_pane: A2       # Freeze row 1 (header stays visible)
freeze_pane: B3       # Freeze row 1-2 AND column A
freeze_row: 1         # Same as freeze_pane: A2
freeze_row: 3         # Freeze first 3 rows
```

### Auto-Filter

```yaml
auto_filter: true     # Filter on the full data range
auto_filter: row 3    # Filter on row 3 specifically
auto_filter: A3:K3    # Explicit range
```

### Formulas

```yaml
formulas:
  # Fill a column with a formula (use {row} as placeholder)
  - column: E
    start_row: 2
    end_row: 100          # Optional, defaults to last row
    formula: "=C{row}+D{row}"

  # Single cell formula
  - cell: B10
    formula: "=SUM(B2:B9)"
```

### Number Formats

```yaml
number_formats:
  # By column
  - column: C
    start_row: 2
    format: "$#,##0.00"

  # By range
  - range: E2:E100
    format: "0.0%"
```

Common formats: `"$#,##0.00"` (currency), `"#,##0"` (thousands), `"0.0%"` (percent), `yyyy-mm-dd` (date)

### Dropdowns

```yaml
# Full format (with validation messages)
dropdowns:
  - range: B2:B100
    values: High,Medium,Low
    allow_blank: true
    prompt: Choose a priority level

# Dict shorthand
dropdowns:
  B: [High, Medium, Low]
  G: [Active, Inactive]
```

### Conditional Formatting

```yaml
# Full format — compare cell values
conditional_formatting:
  - type: cell_is
    range: C2:C100
    operator: greaterThan
    formula: ["10000"]
    style: {fill_color: FF92D050}

  - type: cell_is
    range: F2:F100
    operator: equal
    formula: ['"Blocked"']
    style: {fill_color: FFFF0000, font_color: FFFFFFFF}

# Formula-based — test any condition
  - type: formula
    range: A2:F100
    formula: ['$F2="Active"']
    style: {fill_color: FFDEEBF7}
```

Operators: `equal`, `notEqual`, `greaterThan`, `lessThan`, `greaterThanOrEqual`, `lessThanOrEqual`, `between`, `notBetween`

```yaml
# Legacy shorthand
highlight:
  - when: {G: Blocked}
    fill: red
    font: white
```

### Cell Styles

```yaml
# Dict format
cell_styles:
  A1:D1: header
  A2:D2: subheader

# List format
cell_styles:
  - range: A1:D1
    style: header
```

### Row Styles

Apply a style to entire rows where a column matches a value:
```yaml
row_styles:
  - match: {B: Epic}
    style: highlight_row
```

### Merged Cells

```yaml
merged_cells: [A1:D1, A2:D2]
# or
merge: [A1:D1, A2:D2]
```

Row 1 is auto-merged if only the first cell has content.

## Data Type Coercion

CSV values are automatically converted to proper Excel types:

| CSV value | Excel type |
|-----------|-----------|
| `42` | Number (integer) |
| `3.14` | Number (float) |
| `true` / `false` | Boolean |
| `2026-07-01` | Date |
| `07/01/2026` | Date |
| Everything else | Text |

## Example Files

The `example/` directory contains a ready-to-run workbook:

```
example/
├── workbook.yaml        # Defines 2 sheets with full styling
├── example-sales.csv    # Sales data (formulas, dropdowns, conditional formatting)
└── example-team.csv     # Team directory (merged title, bold, bullet points)
```

Try it:
```bash
cd example
csv-to-xlsx output.xlsx workbook.yaml
```

## Tips

- First row automatically gets the `header` style
- Use `freeze_pane: A2` on every sheet for better scrolling
- Use `auto_filter: true` to enable sorting/filtering in Excel
- Numbers and dates in CSVs are auto-detected, so no manual is conversion needed
- Column letters: A, B, C, ... Z, AA, AB, ...
- Maintain CSV as source of truth and regenerate `.xlsx`
