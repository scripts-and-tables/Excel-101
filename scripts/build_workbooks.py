"""Build all working/ and solutions/ .xlsx files from the source CSV.

Run after `generate_dataset.py`. Re-running is safe — files are overwritten.

Implementation notes:
- Pivot tables cannot be reliably authored via openpyxl, so the solution files
  use SUMIFS-style summary tables that mirror what a learner's pivot would show.
  Each solution sheet that does this includes a top-row note pointing to the
  lesson for the actual pivot build steps.
"""

from __future__ import annotations

import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).resolve().parent.parent
SOURCE_CSV = ROOT / "files" / "source" / "sales_data.csv"
WORKING_DIR = ROOT / "files" / "working"
SOLUTIONS_DIR = ROOT / "files" / "solutions"

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
NOTE_FILL = PatternFill("solid", fgColor="FFF2CC")
NOTE_FONT = Font(italic=True, color="7F6000")


# ----------------------------- helpers --------------------------------------

def load_rows() -> list[dict]:
    with SOURCE_CSV.open(encoding="utf-8") as f:
        reader = csv.DictReader(f)
        rows = []
        for r in reader:
            r["UnitPrice"] = float(r["UnitPrice"])
            r["Quantity"] = int(r["Quantity"])
            r["Discount"] = float(r["Discount"])
            r["Cost"] = float(r["Cost"])
            rows.append(r)
    return rows


HEADERS = ["OrderID", "OrderDate", "Region", "SalesRep", "Customer",
           "Product", "Category", "UnitPrice", "Quantity", "Discount",
           "Cost", "Status"]


def write_sales_sheet(ws, rows: list[dict], *, table_name: str = "Sales",
                      include_revenue: bool = False) -> None:
    """Write the headers + data rows. Optionally add a Revenue formula column."""
    headers = list(HEADERS)
    if include_revenue:
        headers.append("Revenue")

    ws.append(headers)
    for r in rows:
        line = [
            r["OrderID"], r["OrderDate"], r["Region"], r["SalesRep"], r["Customer"],
            r["Product"], r["Category"], r["UnitPrice"], r["Quantity"], r["Discount"],
            r["Cost"], r["Status"],
        ]
        ws.append(line)

    if include_revenue:
        # Add Revenue = UnitPrice * Quantity * (1 - Discount) for every row
        last_row = ws.max_row
        for row_idx in range(2, last_row + 1):
            ws.cell(row=row_idx, column=13,
                    value=f"=H{row_idx}*I{row_idx}*(1-J{row_idx})")

    end_col = get_column_letter(len(headers))
    end_row = ws.max_row
    table_ref = f"A1:{end_col}{end_row}"

    table = Table(displayName=table_name, ref=table_ref)
    style = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    table.tableStyleInfo = style
    ws.add_table(table)

    # Apply nice formatting
    for col_idx, h in enumerate(headers, 1):
        ws.cell(row=1, column=col_idx).font = HEADER_FONT
        ws.cell(row=1, column=col_idx).fill = HEADER_FILL

    # Column widths
    widths = {"A": 12, "B": 12, "C": 11, "D": 18, "E": 24, "F": 22,
              "G": 14, "H": 11, "I": 10, "J": 10, "K": 10, "L": 13, "M": 12}
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w

    # Format Discount as percent, currency for money columns
    for row_idx in range(2, end_row + 1):
        ws.cell(row=row_idx, column=8).number_format = '"$"#,##0.00'   # UnitPrice
        ws.cell(row=row_idx, column=10).number_format = "0%"           # Discount
        ws.cell(row=row_idx, column=11).number_format = '"$"#,##0.00'  # Cost
        if include_revenue:
            ws.cell(row=row_idx, column=13).number_format = '"$"#,##0.00'

    ws.freeze_panes = "A2"


def add_note(ws, text: str, *, cell: str = "A1") -> None:
    ws[cell] = text
    ws[cell].fill = NOTE_FILL
    ws[cell].font = NOTE_FONT
    ws[cell].alignment = Alignment(wrap_text=True, vertical="top")
    # Spread note across a few columns visually
    ws.merge_cells(f"{cell}:H1")
    ws.row_dimensions[1].height = 30


def section(ws, row: int, text: str) -> None:
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(bold=True, color="1F4E78", size=12)


# ------------------------ raw "dirty" data variant --------------------------

def write_dirty_sheet(ws, rows: list[dict]) -> None:
    """For Module 1 working file: keep the dirty rows, no Excel Table yet."""
    ws.append(HEADERS)
    for r in rows:
        ws.append([r[h] for h in HEADERS])

    for col_idx, h in enumerate(HEADERS, 1):
        ws.cell(row=1, column=col_idx).font = HEADER_FONT
        ws.cell(row=1, column=col_idx).fill = HEADER_FILL

    widths = {"A": 12, "B": 12, "C": 11, "D": 18, "E": 24, "F": 22,
              "G": 14, "H": 11, "I": 10, "J": 10, "K": 10, "L": 13}
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w
    ws.freeze_panes = "A2"


# ----------------------------- module 1 -------------------------------------

def build_module_1(rows: list[dict]) -> None:
    # Working: raw, dirty data + an empty Exercises sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "RawData"
    write_dirty_sheet(ws, rows)

    ex = wb.create_sheet("Exercises")
    add_note(ex, "Module 1 exercises — see lesson for instructions.")
    section(ex, 3, "1. Convert RawData into an Excel Table named 'Sales'.")
    section(ex, 4, "2. Trim spaces from Customer names.")
    section(ex, 5, "3. Standardise Region casing (e.g. NORTH -> North).")
    section(ex, 6, "4. Remove exact duplicate orders.")
    section(ex, 7, "5. Multi-sort by Region, then OrderDate descending.")
    wb.save(WORKING_DIR / "module-1.xlsx")

    # Solution: cleaned data as a proper Excel Table
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"

    cleaned = []
    seen: set[tuple] = set()
    for r in rows:
        clean = dict(r)
        clean["Customer"] = clean["Customer"].strip()
        clean["Region"] = clean["Region"].title()  # NORTH -> North
        # Dedupe key excludes OrderID (since dupes have different IDs by design)
        key = (clean["OrderDate"], clean["Region"], clean["SalesRep"],
               clean["Customer"], clean["Product"], clean["Quantity"],
               clean["Discount"], clean["Status"])
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(clean)

    cleaned.sort(key=lambda r: (r["Region"], r["OrderDate"]), reverse=False)
    write_sales_sheet(ws, cleaned, table_name="Sales", include_revenue=False)

    notes = wb.create_sheet("ReadMe", 0)
    add_note(notes,
             "Solution for Module 1. The Sales sheet is now a proper Excel Table "
             "with cleaned customer names, normalised Region casing, no duplicates, "
             "and sorted by Region then OrderDate.")
    wb.save(SOLUTIONS_DIR / "module-1.xlsx")


# ----------------------------- module 2 -------------------------------------

MODULE_2_PROMPTS: list[tuple[str, str, str]] = [
    # (label, working answer cell content, solution formula)
    ("1. Total revenue across all orders", "", "=SUM(Sales[Revenue])"),
    ("2. Total revenue for the West region",
     "", '=SUMIFS(Sales[Revenue],Sales[Region],"West")'),
    ("3. Number of Closed Won orders",
     "", '=COUNTIFS(Sales[Status],"Closed Won")'),
    ("4. Average order revenue for Hardware",
     "", '=AVERAGEIFS(Sales[Revenue],Sales[Category],"Hardware")'),
    ("5. Largest single Software order revenue",
     "", '=MAXIFS(Sales[Revenue],Sales[Category],"Software")'),
    ("6. Lookup the Category for product 'Headset Pro' (XLOOKUP)",
     "", '=XLOOKUP("Headset Pro",Sales[Product],Sales[Category],"not found")'),
    ("7. Tier label: 'Big' if revenue > 5000, 'Mid' if > 1000, else 'Small' (apply to row 2)",
     "", '=IFS(Sales[@Revenue]>5000,"Big",Sales[@Revenue]>1000,"Mid",TRUE,"Small")'),
    ("8. Days since the most recent order",
     "", "=TODAY()-MAX(Sales[OrderDate])"),
    ("9. End of month for the most recent order",
     "", "=EOMONTH(MAX(Sales[OrderDate]),0)"),
    ("10. Unique list of regions (dynamic array)",
     "", "=UNIQUE(Sales[Region])"),
    ("11. Filtered list: all orders for SalesRep 'Anna Becker'",
     "", '=FILTER(Sales,Sales[SalesRep]="Anna Becker","none")'),
    ("12. Initials of SalesRep 'Anna Becker' using LEFT + FIND",
     "", '=LEFT("Anna Becker",1)&MID("Anna Becker",FIND(" ","Anna Becker")+1,1)'),
]


def build_module_2(rows: list[dict]) -> None:
    cleaned = _cleaned_rows(rows)

    for variant, fill_answers in (("working", False), ("solution", True)):
        wb = Workbook()
        ws = wb.active
        ws.title = "Sales"
        write_sales_sheet(ws, cleaned, table_name="Sales", include_revenue=True)

        ex = wb.create_sheet("Exercises")
        add_note(ex,
                 "Module 2 exercises. The Sales sheet uses an Excel Table called 'Sales' "
                 "with a Revenue column = UnitPrice * Quantity * (1 - Discount). "
                 "Put your formula in column C of the matching row.")

        ex["A3"] = "#"
        ex["B3"] = "Question"
        ex["C3"] = "Your formula"
        for c in ("A3", "B3", "C3"):
            ex[c].font = HEADER_FONT
            ex[c].fill = HEADER_FILL
        ex.column_dimensions["A"].width = 4
        ex.column_dimensions["B"].width = 70
        ex.column_dimensions["C"].width = 50

        for i, (label, _, solution) in enumerate(MODULE_2_PROMPTS, start=1):
            row = 3 + i
            ex.cell(row=row, column=1, value=i)
            ex.cell(row=row, column=2, value=label)
            if fill_answers:
                ex.cell(row=row, column=3, value=solution)

        out = (SOLUTIONS_DIR if variant == "solution" else WORKING_DIR) / "module-2.xlsx"
        wb.save(out)


# ----------------------------- module 3 -------------------------------------

def build_module_3(rows: list[dict]) -> None:
    cleaned = _cleaned_rows(rows)

    # Working file: just the data + empty pivot scratchpad
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    write_sales_sheet(ws, cleaned, table_name="Sales", include_revenue=True)

    pad = wb.create_sheet("Pivot Scratchpad")
    add_note(pad, "Build your pivot tables here. See the Module 3 lesson for step-by-step instructions.")
    wb.save(WORKING_DIR / "module-3.xlsx")

    # Solution file: SUMIFS summary tables that mirror the lesson's pivots
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    write_sales_sheet(ws, cleaned, table_name="Sales", include_revenue=True)

    summary = wb.create_sheet("Summary")
    add_note(summary,
             "Solution view: these summary tables use SUMIFS / COUNTIFS to mirror "
             "what your pivot table should show. The lesson covers building the "
             "actual pivot table in Excel.")

    # Revenue by Region
    section(summary, 3, "Revenue by Region")
    summary["A4"] = "Region"; summary["B4"] = "Revenue"
    summary["A4"].font = HEADER_FONT; summary["A4"].fill = HEADER_FILL
    summary["B4"].font = HEADER_FONT; summary["B4"].fill = HEADER_FILL
    regions = sorted({r["Region"].title() for r in cleaned})
    for i, region in enumerate(regions, start=5):
        summary.cell(row=i, column=1, value=region)
        summary.cell(row=i, column=2,
                     value=f'=SUMIFS(Sales[Revenue],Sales[Region],A{i})')
        summary.cell(row=i, column=2).number_format = '"$"#,##0'

    # Revenue by Category and Region (cross-tab)
    section(summary, 13, "Revenue by Category x Region")
    summary["A14"] = "Category"
    summary["A14"].font = HEADER_FONT; summary["A14"].fill = HEADER_FILL
    for j, region in enumerate(regions, start=2):
        cell = summary.cell(row=14, column=j, value=region)
        cell.font = HEADER_FONT; cell.fill = HEADER_FILL
    categories = ["Hardware", "Accessories", "Software", "Services"]
    for i, cat in enumerate(categories, start=15):
        summary.cell(row=i, column=1, value=cat)
        for j, region in enumerate(regions, start=2):
            col_letter = get_column_letter(j)
            summary.cell(row=i, column=j,
                         value=(f'=SUMIFS(Sales[Revenue],Sales[Category],$A{i},'
                                f'Sales[Region],{col_letter}$14)'))
            summary.cell(row=i, column=j).number_format = '"$"#,##0'

    # Top reps by revenue
    section(summary, 22, "Revenue by Sales Rep (sorted manually in solution)")
    rep_totals: dict[str, float] = {}
    for r in cleaned:
        rev = r["UnitPrice"] * r["Quantity"] * (1 - r["Discount"])
        rep_totals[r["SalesRep"]] = rep_totals.get(r["SalesRep"], 0.0) + rev
    rep_rows = sorted(rep_totals.items(), key=lambda kv: -kv[1])
    summary["A23"] = "SalesRep"; summary["B23"] = "Revenue"
    for c in ("A23", "B23"):
        summary[c].font = HEADER_FONT; summary[c].fill = HEADER_FILL
    for i, (rep, total) in enumerate(rep_rows, start=24):
        summary.cell(row=i, column=1, value=rep)
        summary.cell(row=i, column=2,
                     value=f'=SUMIFS(Sales[Revenue],Sales[SalesRep],A{i})')
        summary.cell(row=i, column=2).number_format = '"$"#,##0'

    summary.column_dimensions["A"].width = 22
    for col in "BCDEFG":
        summary.column_dimensions[col].width = 16

    # Add a small bar chart of revenue by region
    chart = BarChart()
    chart.type = "col"
    chart.title = "Revenue by Region"
    chart.y_axis.title = "Revenue"
    chart.x_axis.title = "Region"
    data = Reference(summary, min_col=2, min_row=4, max_row=4 + len(regions))
    cats = Reference(summary, min_col=1, min_row=5, max_row=4 + len(regions))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 8
    chart.width = 16
    summary.add_chart(chart, "D3")

    wb.save(SOLUTIONS_DIR / "module-3.xlsx")


# ----------------------------- module 4 -------------------------------------

def build_module_4(rows: list[dict]) -> None:
    cleaned = _cleaned_rows(rows)

    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    write_sales_sheet(ws, cleaned, table_name="Sales", include_revenue=True)
    pad = wb.create_sheet("Charts")
    add_note(pad, "Build the charts described in the Module 4 lesson on this sheet.")
    wb.save(WORKING_DIR / "module-4.xlsx")

    # Solution: monthly revenue + chart, conditional formatting on region totals
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    write_sales_sheet(ws, cleaned, table_name="Sales", include_revenue=True)

    monthly = wb.create_sheet("Monthly")
    add_note(monthly, "Monthly revenue trend with line chart (Module 4 solution).")
    monthly["A3"] = "Month"; monthly["B3"] = "Revenue"
    monthly["A3"].font = HEADER_FONT; monthly["A3"].fill = HEADER_FILL
    monthly["B3"].font = HEADER_FONT; monthly["B3"].fill = HEADER_FILL

    by_month: dict[str, float] = {}
    for r in cleaned:
        month_key = r["OrderDate"][:7]  # YYYY-MM
        rev = r["UnitPrice"] * r["Quantity"] * (1 - r["Discount"])
        by_month[month_key] = by_month.get(month_key, 0.0) + rev
    for i, month in enumerate(sorted(by_month.keys()), start=4):
        monthly.cell(row=i, column=1, value=month)
        monthly.cell(row=i, column=2, value=round(by_month[month], 2))
        monthly.cell(row=i, column=2).number_format = '"$"#,##0'
    monthly.column_dimensions["A"].width = 12
    monthly.column_dimensions["B"].width = 16

    last = monthly.max_row
    line = LineChart()
    line.title = "Monthly Revenue Trend"
    line.y_axis.title = "Revenue"
    line.x_axis.title = "Month"
    data = Reference(monthly, min_col=2, min_row=3, max_row=last)
    cats = Reference(monthly, min_col=1, min_row=4, max_row=last)
    line.add_data(data, titles_from_data=True)
    line.set_categories(cats)
    line.height = 9
    line.width = 18
    monthly.add_chart(line, "D3")

    # Conditional formatting demo on a region table
    cf = wb.create_sheet("Heatmap")
    add_note(cf, "Conditional formatting demo: revenue per region per quarter, with a colour scale.")
    cf["A3"] = "Region"
    cf["A3"].font = HEADER_FONT; cf["A3"].fill = HEADER_FILL
    quarters = ["2024-Q1", "2024-Q2", "2024-Q3", "2024-Q4",
                "2025-Q1", "2025-Q2", "2025-Q3", "2025-Q4"]
    for j, q in enumerate(quarters, start=2):
        c = cf.cell(row=3, column=j, value=q)
        c.font = HEADER_FONT; c.fill = HEADER_FILL

    regions = sorted({r["Region"].title() for r in cleaned})
    region_q: dict[tuple[str, str], float] = {}
    for r in cleaned:
        year = r["OrderDate"][:4]
        month = int(r["OrderDate"][5:7])
        q = f"{year}-Q{(month - 1) // 3 + 1}"
        rev = r["UnitPrice"] * r["Quantity"] * (1 - r["Discount"])
        region_q[(r["Region"].title(), q)] = region_q.get((r["Region"].title(), q), 0.0) + rev

    for i, region in enumerate(regions, start=4):
        cf.cell(row=i, column=1, value=region)
        for j, q in enumerate(quarters, start=2):
            v = round(region_q.get((region, q), 0.0), 2)
            cell = cf.cell(row=i, column=j, value=v)
            cell.number_format = '"$"#,##0'
    cf.column_dimensions["A"].width = 14
    for j in range(2, 2 + len(quarters)):
        cf.column_dimensions[get_column_letter(j)].width = 12

    end_col = get_column_letter(1 + len(quarters))
    end_row = 3 + len(regions)
    cf.conditional_formatting.add(
        f"B4:{end_col}{end_row}",
        ColorScaleRule(start_type="min", start_color="FFFFFF",
                       mid_type="percentile", mid_value=50, mid_color="FFEB84",
                       end_type="max", end_color="63BE7B"),
    )

    # Data bars on rep totals
    bars = wb.create_sheet("RepBars")
    add_note(bars, "Data bars on Sales Rep totals — quick visual ranking.")
    bars["A3"] = "SalesRep"; bars["B3"] = "Revenue"
    bars["A3"].font = HEADER_FONT; bars["A3"].fill = HEADER_FILL
    bars["B3"].font = HEADER_FONT; bars["B3"].fill = HEADER_FILL
    rep_totals: dict[str, float] = {}
    for r in cleaned:
        rev = r["UnitPrice"] * r["Quantity"] * (1 - r["Discount"])
        rep_totals[r["SalesRep"]] = rep_totals.get(r["SalesRep"], 0.0) + rev
    for i, (rep, total) in enumerate(sorted(rep_totals.items(), key=lambda kv: -kv[1]), start=4):
        bars.cell(row=i, column=1, value=rep)
        bars.cell(row=i, column=2, value=round(total, 2))
        bars.cell(row=i, column=2).number_format = '"$"#,##0'
    end_row = 3 + len(rep_totals)
    bars.conditional_formatting.add(
        f"B4:B{end_row}",
        DataBarRule(start_type="min", end_type="max", color="638EC6"),
    )
    bars.column_dimensions["A"].width = 22
    bars.column_dimensions["B"].width = 16

    wb.save(SOLUTIONS_DIR / "module-4.xlsx")


# ----------------------------- module 5 -------------------------------------

def build_module_5(rows: list[dict]) -> None:
    cleaned = _cleaned_rows(rows)

    # Working: data + empty calculator stub
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    write_sales_sheet(ws, cleaned, table_name="Sales", include_revenue=True)
    calc = wb.create_sheet("Quote Calculator")
    add_note(calc, "Build the quote calculator described in the Module 5 lesson.")
    wb.save(WORKING_DIR / "module-5.xlsx")

    # Solution: full calculator with named ranges, validation, Goal Seek-ready cell
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    write_sales_sheet(ws, cleaned, table_name="Sales", include_revenue=True)

    calc = wb.create_sheet("Quote Calculator")
    add_note(calc, "Quote calculator with data validation drop-downs and named ranges. "
                   "Use Goal Seek (Data > What-If) on cell B12 to find the discount that hits a target margin.")

    calc["A3"] = "Region";    calc["B3"] = "West"
    calc["A4"] = "Product";   calc["B4"] = "Laptop Pro 14"
    calc["A5"] = "Quantity";  calc["B5"] = 10
    calc["A6"] = "Discount";  calc["B6"] = 0.10
    calc["B6"].number_format = "0%"

    calc["A8"] = "Unit Price"
    calc["B8"] = '=XLOOKUP(B4,Sales[Product],Sales[UnitPrice],"not found")'
    calc["B8"].number_format = '"$"#,##0.00'

    calc["A9"] = "Unit Cost"
    calc["B9"] = '=XLOOKUP(B4,Sales[Product],Sales[Cost],"not found")'
    calc["B9"].number_format = '"$"#,##0.00'

    calc["A10"] = "Revenue"
    calc["B10"] = "=B8*B5*(1-B6)"
    calc["B10"].number_format = '"$"#,##0.00'

    calc["A11"] = "Total Cost"
    calc["B11"] = "=B9*B5"
    calc["B11"].number_format = '"$"#,##0.00'

    calc["A12"] = "Margin %"
    calc["B12"] = "=IFERROR((B10-B11)/B10,0)"
    calc["B12"].number_format = "0.0%"

    for r in (3, 4, 5, 6, 8, 9, 10, 11, 12):
        calc.cell(row=r, column=1).font = Font(bold=True)
    calc.column_dimensions["A"].width = 16
    calc.column_dimensions["B"].width = 22

    # Data validation: Region drop-down
    region_list = '"Central,East,North,South,West"'
    dv_region = DataValidation(type="list", formula1=region_list, allow_blank=False)
    dv_region.add("B3")
    calc.add_data_validation(dv_region)

    # Data validation: Product list — point at a helper sheet
    helper = wb.create_sheet("Lists")
    helper["A1"] = "Products"
    helper["A1"].font = HEADER_FONT
    helper["A1"].fill = HEADER_FILL
    products = sorted({r["Product"] for r in cleaned})
    for i, p in enumerate(products, start=2):
        helper.cell(row=i, column=1, value=p)
    helper.column_dimensions["A"].width = 24
    last_p = 1 + len(products)

    dv_product = DataValidation(type="list", formula1=f"=Lists!$A$2:$A${last_p}", allow_blank=False)
    dv_product.add("B4")
    calc.add_data_validation(dv_product)

    # Named ranges for clarity
    wb.defined_names["ProductList"] = DefinedName(
        name="ProductList", attr_text=f"Lists!$A$2:$A${last_p}")
    wb.defined_names["SelectedProduct"] = DefinedName(
        name="SelectedProduct", attr_text="'Quote Calculator'!$B$4")

    wb.save(SOLUTIONS_DIR / "module-5.xlsx")


# ----------------------------- capstone -------------------------------------

def build_capstone(rows: list[dict]) -> None:
    cleaned = _cleaned_rows(rows)

    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    write_sales_sheet(ws, cleaned, table_name="Sales", include_revenue=True)
    brief = wb.create_sheet("Brief", 0)
    add_note(brief, "Capstone brief — see the Module 5 lesson for the full task. "
                    "Build a one-page dashboard on a new sheet called 'Dashboard'.")
    wb.save(WORKING_DIR / "capstone.xlsx")

    # Solution capstone with KPI cards + summary tables + charts
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    write_sales_sheet(ws, cleaned, table_name="Sales", include_revenue=True)

    dash = wb.create_sheet("Dashboard", 0)
    add_note(dash, "Capstone solution dashboard.")

    # KPIs
    kpi_specs = [
        ("Total Revenue (Closed Won)", '=SUMIFS(Sales[Revenue],Sales[Status],"Closed Won")', '"$"#,##0'),
        ("Total Orders (Closed Won)",  '=COUNTIFS(Sales[Status],"Closed Won")', "#,##0"),
        ("Avg Order Value",            '=IFERROR(SUMIFS(Sales[Revenue],Sales[Status],"Closed Won")/COUNTIFS(Sales[Status],"Closed Won"),0)', '"$"#,##0'),
        ("Refund Rate",                '=IFERROR(COUNTIFS(Sales[Status],"Refunded")/COUNTA(Sales[OrderID]),0)', "0.0%"),
    ]
    for i, (label, formula, fmt) in enumerate(kpi_specs):
        col = 1 + i * 2
        dash.cell(row=3, column=col, value=label).font = Font(bold=True, color="1F4E78")
        cell = dash.cell(row=4, column=col, value=formula)
        cell.font = Font(bold=True, size=14)
        cell.number_format = fmt
        dash.column_dimensions[get_column_letter(col)].width = 22

    # Revenue by region (formula-backed)
    section(dash, 7, "Revenue by Region")
    dash["A8"] = "Region"; dash["B8"] = "Revenue"
    dash["A8"].font = HEADER_FONT; dash["A8"].fill = HEADER_FILL
    dash["B8"].font = HEADER_FONT; dash["B8"].fill = HEADER_FILL
    regions = sorted({r["Region"].title() for r in cleaned})
    for i, region in enumerate(regions, start=9):
        dash.cell(row=i, column=1, value=region)
        dash.cell(row=i, column=2,
                  value=f'=SUMIFS(Sales[Revenue],Sales[Region],A{i})')
        dash.cell(row=i, column=2).number_format = '"$"#,##0'
    region_chart = BarChart()
    region_chart.type = "bar"
    region_chart.title = "Revenue by Region"
    data = Reference(dash, min_col=2, min_row=8, max_row=8 + len(regions))
    cats = Reference(dash, min_col=1, min_row=9, max_row=8 + len(regions))
    region_chart.add_data(data, titles_from_data=True)
    region_chart.set_categories(cats)
    region_chart.height = 7
    region_chart.width = 14
    dash.add_chart(region_chart, "D7")

    # Monthly trend
    section(dash, 22, "Monthly Revenue")
    dash["A23"] = "Month"; dash["B23"] = "Revenue"
    dash["A23"].font = HEADER_FONT; dash["A23"].fill = HEADER_FILL
    dash["B23"].font = HEADER_FONT; dash["B23"].fill = HEADER_FILL
    by_month: dict[str, float] = {}
    for r in cleaned:
        rev = r["UnitPrice"] * r["Quantity"] * (1 - r["Discount"])
        by_month[r["OrderDate"][:7]] = by_month.get(r["OrderDate"][:7], 0.0) + rev
    for i, month in enumerate(sorted(by_month.keys()), start=24):
        dash.cell(row=i, column=1, value=month)
        dash.cell(row=i, column=2, value=round(by_month[month], 2))
        dash.cell(row=i, column=2).number_format = '"$"#,##0'
    last = dash.max_row
    line = LineChart()
    line.title = "Monthly Revenue"
    data = Reference(dash, min_col=2, min_row=23, max_row=last)
    cats = Reference(dash, min_col=1, min_row=24, max_row=last)
    line.add_data(data, titles_from_data=True)
    line.set_categories(cats)
    line.height = 8
    line.width = 18
    dash.add_chart(line, "D22")

    dash.column_dimensions["A"].width = 14
    dash.column_dimensions["B"].width = 16

    wb.save(SOLUTIONS_DIR / "capstone.xlsx")


# ----------------------------- shared helpers --------------------------------

def _cleaned_rows(rows: list[dict]) -> list[dict]:
    cleaned = []
    seen = set()
    for r in rows:
        c = dict(r)
        c["Customer"] = c["Customer"].strip()
        c["Region"] = c["Region"].title()
        key = (c["OrderDate"], c["Region"], c["SalesRep"], c["Customer"],
               c["Product"], c["Quantity"], c["Discount"], c["Status"])
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(c)
    return cleaned


# ------------------------------ entry ---------------------------------------

def main() -> None:
    WORKING_DIR.mkdir(parents=True, exist_ok=True)
    SOLUTIONS_DIR.mkdir(parents=True, exist_ok=True)

    rows = load_rows()
    print(f"Loaded {len(rows)} source rows")

    build_module_1(rows);    print("  module 1 done")
    build_module_2(rows);    print("  module 2 done")
    build_module_3(rows);    print("  module 3 done")
    build_module_4(rows);    print("  module 4 done")
    build_module_5(rows);    print("  module 5 done")
    build_capstone(rows);    print("  capstone done")


if __name__ == "__main__":
    main()
