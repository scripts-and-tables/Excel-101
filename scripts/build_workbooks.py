"""Build all working/ and solutions/ .xlsx files from the source CSVs.

Run after `generate_dataset.py`. Re-running is safe — files are overwritten.

The course follows the real sales-data workflow on one running dataset. Each
module is a STAGE of that process (file names stay module-N.xlsx, mapped 1:1):
- module-1  Stage 1 · Receive & clean   (Table, TRIM/CLEAN, Find&Replace, VALUE/DATEVALUE,
                                          text functions, dedupe, formatting)
- module-2  Stage 2 · Structure & enrich (sort/filter; VLOOKUP manager/quota/commission rate
                                          from the Reps table; failure modes; XLOOKUP note)
- module-3  Stage 3 · Explore & summarize (AutoFilter, Status Bar, SUBTOTAL; SUMIFS/COUNTIFS/
                                          AVERAGEIFS; sales KPIs; commission calc)
- module-4  Stage 4 · Pivot & rank        (PivotTables/slicers, mirrored with SUMIFS; leaderboard)
- module-5  Stage 5 · Present             (KPI one-pager, one chart, conditional formatting)
- capstone  fresh file: run the whole workflow and answer ~11 questions.

Implementation notes:
- PivotTables, Slicers and Data > Subtotal outlines cannot be reliably authored
  via openpyxl, so the solution files use SUMIFS / SUBTOTAL summary tables that
  mirror what a learner's pivot or subtotal would show. Each such sheet carries a
  top-row note pointing back to the lesson for the real build steps.
"""

from __future__ import annotations

import csv
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parent.parent
# Files live under docs/ so GitHub Pages serves them for download.
FILES = ROOT / "docs" / "files"
SOURCE_CSV = FILES / "source" / "sales_data.csv"
REPS_CSV = FILES / "source" / "reps.csv"
WORKING_DIR = FILES / "working"
SOLUTIONS_DIR = FILES / "solutions"

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
NOTE_FILL = PatternFill("solid", fgColor="FFF2CC")
NOTE_FONT = Font(italic=True, color="7F6000")
TASK_FILL = PatternFill("solid", fgColor="E2EFDA")


# ----------------------------- data loading ---------------------------------

HEADERS = ["OrderID", "OrderDate", "Region", "SalesRep", "Customer",
           "Product", "Category", "UnitPrice", "Quantity", "Discount",
           "Cost", "Status"]


def load_rows() -> list[dict]:
    with SOURCE_CSV.open(encoding="utf-8") as f:
        rows = []
        for r in csv.DictReader(f):
            r["UnitPrice"] = float(r["UnitPrice"])
            r["Quantity"] = int(r["Quantity"])
            r["Discount"] = float(r["Discount"])
            r["Cost"] = float(r["Cost"])
            rows.append(r)
    return rows


def load_reps() -> list[dict]:
    with REPS_CSV.open(encoding="utf-8") as f:
        reps = []
        for r in csv.DictReader(f):
            r["AnnualQuota"] = int(r["AnnualQuota"])
            r["CommissionRate"] = float(r["CommissionRate"])
            reps.append(r)
    return reps


def _as_date(iso: str) -> date:
    y, m, d = iso.split("-")
    return date(int(y), int(m), int(d))


def _cleaned_rows(rows: list[dict]) -> list[dict]:
    """Trim customer names, normalise Region casing, drop exact-duplicate orders."""
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


# ----------------------------- sheet writers --------------------------------

COL_WIDTHS = {"A": 12, "B": 12, "C": 11, "D": 18, "E": 24, "F": 22,
              "G": 14, "H": 11, "I": 10, "J": 10, "K": 11, "L": 13, "M": 13}


def _style_header(ws, ncols: int) -> None:
    for col_idx in range(1, ncols + 1):
        ws.cell(row=1, column=col_idx).font = HEADER_FONT
        ws.cell(row=1, column=col_idx).fill = HEADER_FILL
    for col_letter, w in COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = w


def _row_values(r: dict, *, typed_dates: bool) -> list:
    return [
        r["OrderID"],
        _as_date(r["OrderDate"]) if typed_dates else r["OrderDate"],
        r["Region"], r["SalesRep"], r["Customer"], r["Product"], r["Category"],
        r["UnitPrice"], r["Quantity"], r["Discount"], r["Cost"], r["Status"],
    ]


def _format_data_rows(ws, last_row: int, *, has_revenue: bool) -> None:
    for row_idx in range(2, last_row + 1):
        ws.cell(row=row_idx, column=2).number_format = "yyyy-mm-dd"   # OrderDate
        ws.cell(row=row_idx, column=8).number_format = '"$"#,##0.00'  # UnitPrice
        ws.cell(row=row_idx, column=10).number_format = "0%"          # Discount
        ws.cell(row=row_idx, column=11).number_format = '"$"#,##0.00' # Cost
        if has_revenue:
            ws.cell(row=row_idx, column=13).number_format = '"$"#,##0.00'


def write_sales_table(ws, rows: list[dict], *, table_name: str = "Sales",
                      include_revenue: bool = False) -> None:
    """Clean data as a proper Excel Table, with optional Revenue column."""
    headers = list(HEADERS) + (["Revenue"] if include_revenue else [])
    ws.append(headers)
    for r in rows:
        ws.append(_row_values(r, typed_dates=True))

    if include_revenue:
        for row_idx in range(2, ws.max_row + 1):
            ws.cell(row=row_idx, column=13,
                    value=f"=H{row_idx}*I{row_idx}*(1-J{row_idx})")

    end_col = get_column_letter(len(headers))
    table = Table(displayName=table_name, ref=f"A1:{end_col}{ws.max_row}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)

    _style_header(ws, len(headers))
    _format_data_rows(ws, ws.max_row, has_revenue=include_revenue)
    ws.freeze_panes = "A2"


def write_sales_range(ws, rows: list[dict], *, include_revenue: bool = False) -> None:
    """Plain range (no Excel Table)."""
    headers = list(HEADERS) + (["Revenue"] if include_revenue else [])
    ws.append(headers)
    for r in rows:
        vals = _row_values(r, typed_dates=True)
        if include_revenue:
            vals.append(round(r["UnitPrice"] * r["Quantity"] * (1 - r["Discount"]), 2))
        ws.append(vals)
    _style_header(ws, len(headers))
    _format_data_rows(ws, ws.max_row, has_revenue=include_revenue)
    ws.freeze_panes = "A2"


def write_reps_table(ws, reps: list[dict], *, table_name: str = "Reps") -> None:
    """Sales Rep reference table: rep -> region, manager, quota, commission rate.
    Lookup target for Stage 2 (enrich) and the Stage 3 commission calc."""
    ws.append(["SalesRep", "Region", "Manager", "AnnualQuota", "CommissionRate"])
    for r in reps:
        ws.append([r["SalesRep"], r["Region"], r["Manager"],
                   r["AnnualQuota"], r["CommissionRate"]])
    table = Table(displayName=table_name, ref=f"A1:E{ws.max_row}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium4", showRowStripes=True)
    ws.add_table(table)
    for col_idx in range(1, 6):
        ws.cell(row=1, column=col_idx).font = HEADER_FONT
        ws.cell(row=1, column=col_idx).fill = HEADER_FILL
    for row_idx in range(2, ws.max_row + 1):
        ws.cell(row=row_idx, column=4).number_format = '"$"#,##0'
        ws.cell(row=row_idx, column=5).number_format = "0%"
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 11
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 15
    ws.freeze_panes = "A2"


# ----------------------------- small helpers --------------------------------

def add_note(ws, text: str, *, cell: str = "A1", span: str = "H") -> None:
    ws[cell] = text
    ws[cell].fill = NOTE_FILL
    ws[cell].font = NOTE_FONT
    ws[cell].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(f"{cell}:{span}1")
    ws.row_dimensions[1].height = 30


def section(ws, row: int, text: str) -> None:
    ws.cell(row=row, column=1, value=text).font = Font(bold=True, color="1F4E78", size=12)


def task_header(ws, row: int, cols: list[str], widths: list[int]) -> None:
    for i, (label, w) in enumerate(zip(cols, widths), start=1):
        c = ws.cell(row=row, column=i, value=label)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        ws.column_dimensions[get_column_letter(i)].width = w


def hint(ws, row: int, text: str) -> None:
    ws.cell(row=row, column=2, value=text).font = Font(italic=True, color="808080")


def task_list(ws, items: list[str], *, start: int = 3) -> int:
    for i, t in enumerate(items, start=start):
        section(ws, i, t)
    return start + len(items)


# ===================== Stage 1 · Receive & clean ============================

def build_stage1(rows: list[dict]) -> None:
    # ---- working ----
    wb = Workbook()
    ws = wb.active
    ws.title = "RawData"
    write_sales_range(ws, rows, include_revenue=False)
    add_note(ws, "RAW export — messy on purpose. Step 1 of the workflow: turn this into a "
                 "clean Excel Table called 'Sales'. Watch for spaces in Customer names, "
                 "MiXeD-case Regions, and duplicate orders.", cell="N1", span="N")
    ws.column_dimensions["N"].width = 4

    # focused scratch sheet: cleaning functions (conversions + text)
    clean = wb.create_sheet("Cleanup Practice")
    add_note(clean, "Cleanup Practice — fix each messy value in column B using the function "
                    "in column A; type your formula in column C.")
    task_header(clean, 3, ["Function", "Messy value", "Your fix"], [16, 30, 30])
    # kind: "text" = stored as text (so VALUE/DATEVALUE/TRIM/text fns have a job);
    #       "formula" = a CHAR()-built value with a hidden non-printable char for CLEAN.
    practice = [
        ("TRIM", "  Acme  Corp  ", "text", '=TRIM(B4)'),
        ("CLEAN", '=CHAR(10)&"Globex Industries"', "formula", '=CLEAN(B5)'),
        ("TRIM+CLEAN", '=CHAR(9)&"  Initech LLC  "', "formula", '=TRIM(CLEAN(B6))'),
        ("VALUE", "1250", "text", '=VALUE(B7)'),
        ("VALUE", "89.50", "text", '=VALUE(B8)'),
        ("DATEVALUE", "2024-03-15", "text", '=DATEVALUE(B9)'),
        ("DATEVALUE", "2025-11-02", "text", '=DATEVALUE(B10)'),
        ("SUBSTITUTE", "Acme;Corp;LLC", "text", '=SUBSTITUTE(B11,";"," ")'),
        ("LEFT", "ORD-2024-0042", "text", '=LEFT(B12,3)'),
        ("MID", "ORD-2024-0042", "text", '=MID(B13,5,4)'),
        ("RIGHT", "ORD-2024-0042", "text", '=RIGHT(B14,4)'),
        ('& (combine)', "0042", "text", '="ORD-"&B15'),
    ]
    _write_practice(clean, practice, fill=False)
    tip_row = 4 + len(practice) + 1
    clean.cell(row=tip_row, column=1, value="Tip:").font = Font(bold=True, color="7F6000")
    hint(clean, tip_row, "Text-stored numbers/dates sit on the LEFT of the cell; real ones sit on the "
                         "right. VALUE / DATEVALUE convert them. LEFT/MID/RIGHT split an ID; & joins text.")

    ex = wb.create_sheet("Exercises")
    add_note(ex, "Stage 1 exercises — clean the raw export. See the lesson for full steps.")
    task_list(ex, [
        "1. Select the RawData range and press Ctrl+T to make an Excel Table; name it 'Sales'.",
        "2. TRIM the spaces from Customer names (helper column, then paste values back).",
        "3. CLEAN any hidden line breaks/tabs (see the Cleanup Practice sheet).",
        "4. Find & Replace (Ctrl+H): fix the CAPS regions — replace 'NORTH' with 'North', etc.",
        "5. Convert the text numbers/dates on Cleanup Practice with VALUE and DATEVALUE.",
        "6. Use LEFT/MID/RIGHT and SUBSTITUTE to split / tidy an order ID; combine text with &.",
        "7. Paste Special > Values to lock a helper column, then delete the original.",
        "8. Data > Remove Duplicates on the order columns — how many were removed?",
        "9. Multi-sort: Region (A-Z), then OrderDate (newest first); freeze the header row.",
        "10. Format: UnitPrice & Cost as currency, Discount as %, and try the $1.2M custom format.",
    ])
    section(ex, 15, "Custom number format to try (Ctrl+1 > Custom):")
    ex.cell(row=16, column=1,
            value='[>=1000000]"$"#,##0,,"M";[>=1000]"$"#,##0,"K";"$"#,##0').font = Font(name="Consolas")
    wb.save(WORKING_DIR / "module-1.xlsx")

    # ---- solution ----
    wb = Workbook()
    rm = wb.active
    rm.title = "ReadMe"
    add_note(rm, "Stage 1 solution. 'Sales' is a clean Excel Table: trimmed names, normalised "
                 "Region casing, duplicates removed, sorted by Region then OrderDate. "
                 "'Cleanup Practice' shows the TRIM/CLEAN/VALUE/DATEVALUE/text-function fixes.")
    ws = wb.create_sheet("Sales")
    cleaned = sorted(_cleaned_rows(rows), key=lambda r: (r["Region"], r["OrderDate"]))
    write_sales_table(ws, cleaned, table_name="Sales", include_revenue=False)

    clean = wb.create_sheet("Cleanup Practice")
    add_note(clean, "Cleanup Practice — solved.")
    task_header(clean, 3, ["Function", "Messy value", "Fixed"], [16, 30, 30])
    _write_practice(clean, practice, fill=True)
    wb.save(SOLUTIONS_DIR / "module-1.xlsx")


def _write_practice(ws, practice, *, fill: bool) -> None:
    for i, (fn, messy, kind, sol) in enumerate(practice, start=4):
        ws.cell(row=i, column=1, value=fn).font = Font(bold=True)
        c = ws.cell(row=i, column=2)
        c.value = messy
        if kind == "text":
            c.data_type = "s"
        if fill:
            out = ws.cell(row=i, column=3, value=sol)
            if fn == "DATEVALUE":
                out.number_format = "yyyy-mm-dd"
            elif fn == "VALUE":
                out.number_format = "#,##0.00"


# ===================== Stage 2 · Structure & enrich =========================
# Sort/filter into shape; VLOOKUP manager / quota / commission rate from Reps.

def build_stage2(rows: list[dict], reps: list[dict]) -> None:
    cleaned = _cleaned_rows(rows)

    # ---- working ----
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    write_sales_table(ws, cleaned, table_name="Sales", include_revenue=True)
    rs = wb.create_sheet("Reps")
    write_reps_table(rs, reps, table_name="Reps")
    add_note(rs, "Reference table. The export doesn't carry a rep's Manager, Quota or "
                 "Commission Rate — VLOOKUP them in from here.", cell="G1", span="J")

    ex = wb.create_sheet("Exercises", 0)
    add_note(ex, "Stage 2 exercises — enrich 'Sales' with VLOOKUP from the 'Reps' table.")
    task_list(ex, [
        "1. Add a 'Manager' column:    =VLOOKUP([@SalesRep], Reps, 3, FALSE).",
        "2. Add a 'Quota' column:      =VLOOKUP([@SalesRep], Reps, 4, FALSE).",
        "3. Add a 'Comm Rate' column:  =VLOOKUP([@SalesRep], Reps, 5, FALSE).  (you'll use this in Stage 3)",
        "4. Why FALSE? Try TRUE on a rep and watch the answer go wrong.",
        "5. VLOOKUP can't look LEFT: try to fetch SalesRep from Manager — what happens?",
        "6. Wrap a missing-rep lookup in IFERROR(..., \"unknown\") to kill the #N/A.",
        "7. (M365) Redo task 1 with XLOOKUP: =XLOOKUP([@SalesRep], Reps[SalesRep], Reps[Manager], \"unknown\").",
    ])
    wb.save(WORKING_DIR / "module-2.xlsx")

    # ---- solution ----
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    headers = list(HEADERS) + ["Revenue", "Manager", "Quota", "Comm Rate"]
    ws.append(headers)
    for r in cleaned:
        ws.append(_row_values(r, typed_dates=True))
    for row_idx in range(2, ws.max_row + 1):
        ws.cell(row=row_idx, column=13, value=f"=H{row_idx}*I{row_idx}*(1-J{row_idx})")
        ws.cell(row=row_idx, column=14, value=f"=VLOOKUP(D{row_idx},Reps,3,FALSE)")
        ws.cell(row=row_idx, column=15, value=f"=VLOOKUP(D{row_idx},Reps,4,FALSE)")
        ws.cell(row=row_idx, column=15).number_format = '"$"#,##0'
        ws.cell(row=row_idx, column=16, value=f"=VLOOKUP(D{row_idx},Reps,5,FALSE)")
        ws.cell(row=row_idx, column=16).number_format = "0%"
    end_col = get_column_letter(len(headers))
    table = Table(displayName="Sales", ref=f"A1:{end_col}{ws.max_row}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)
    _style_header(ws, len(headers))
    ws.column_dimensions["N"].width = 16
    ws.column_dimensions["O"].width = 12
    ws.column_dimensions["P"].width = 11
    _format_data_rows(ws, ws.max_row, has_revenue=True)
    ws.freeze_panes = "A2"

    rs = wb.create_sheet("Reps")
    write_reps_table(rs, reps, table_name="Reps")

    fm = wb.create_sheet("Failure Modes", 0)
    add_note(fm, "The six classic VLOOKUP failures — and the fix for each. (Microsoft "
                 "VLOOKUP troubleshooting.)")
    task_header(fm, 3, ["#", "Failure mode", "Fix"], [4, 46, 60])
    failures = [
        ("Approximate match returns wrong row", "Always pass FALSE (exact) as the 4th argument."),
        ("Can't look to the LEFT of the key column",
         "Re-order columns so the key is leftmost, or use INDEX/MATCH / XLOOKUP."),
        ("Column index breaks when a column is inserted",
         "Reference a Table (Reps) so the index tracks, or use XLOOKUP."),
        ("Number-stored-as-text vs real number mismatch",
         "Make both sides the same type (VALUE, or Text-to-Columns)."),
        ("Trailing/leading spaces in the key",
         "TRIM both the lookup value and the key column first (Stage 1!)."),
        ("#N/A when the value genuinely isn't there",
         "Wrap in IFERROR(VLOOKUP(...), \"unknown\")."),
    ]
    for i, (mode, fix) in enumerate(failures, start=4):
        fm.cell(row=i, column=1, value=i - 3)
        fm.cell(row=i, column=2, value=mode)
        fm.cell(row=i, column=3, value=fix)
    section(fm, 12, "Modern note (Excel 365 / 2021+)")
    fm["B13"] = "XLOOKUP fixes most of these by default — exact match, looks any direction, " \
                "survives column inserts, built-in if-not-found:"
    fm["B14"] = '=XLOOKUP([@SalesRep], Reps[SalesRep], Reps[Manager], "unknown")'
    fm["B14"].font = Font(name="Consolas")
    wb.save(SOLUTIONS_DIR / "module-2.xlsx")


# ===================== Stage 3 · Explore & summarize ========================
# AutoFilter / Status Bar / SUBTOTAL; SUMIFS/COUNTIFS/AVERAGEIFS; KPIs; commission.

STAGE3_PROMPTS: list[tuple[str, str]] = [
    ("1. Total revenue across all orders", "=SUM(Sales[Revenue])"),
    ("2. Number of orders (count OrderIDs)", "=COUNTA(Sales[OrderID])"),
    ("3. Number of Closed Won orders", '=COUNTIFS(Sales[Status],"Closed Won")'),
    ("4. Total revenue for the West region", '=SUMIFS(Sales[Revenue],Sales[Region],"West")'),
    ("5. Closed Won revenue for the West region",
     '=SUMIFS(Sales[Revenue],Sales[Region],"West",Sales[Status],"Closed Won")'),
    ("6. Average order revenue for the Hardware category",
     '=AVERAGEIFS(Sales[Revenue],Sales[Category],"Hardware")'),
    ("7. How many Hardware orders are Refunded?",
     '=COUNTIFS(Sales[Category],"Hardware",Sales[Status],"Refunded")'),
    ("8. KPI — Win rate (Closed Won / all orders)",
     '=COUNTIFS(Sales[Status],"Closed Won")/COUNTA(Sales[OrderID])'),
    ("9. KPI — Average Order Value (Closed Won revenue / Closed Won orders)",
     '=SUMIFS(Sales[Revenue],Sales[Status],"Closed Won")/COUNTIFS(Sales[Status],"Closed Won")'),
    ("10. Commission for Anna Becker (her Closed Won revenue x her rate from Reps)",
     '=SUMIFS(Sales[Revenue],Sales[SalesRep],"Anna Becker",Sales[Status],"Closed Won")'
     '*VLOOKUP("Anna Becker",Reps,5,FALSE)'),
    ("11. Row label: 'Large' if Revenue > 2000, else 'Standard' (helper column)",
     '=IF([@Revenue]>2000,"Large","Standard")'),
]


def build_stage3(rows: list[dict], reps: list[dict]) -> None:
    cleaned = _cleaned_rows(rows)
    for variant, fill in (("working", False), ("solution", True)):
        wb = Workbook()
        ws = wb.active
        ws.title = "Sales"
        write_sales_table(ws, cleaned, table_name="Sales", include_revenue=True)
        rs = wb.create_sheet("Reps")
        write_reps_table(rs, reps, table_name="Reps")

        ex = wb.create_sheet("Exercises", 0)
        add_note(ex, "Stage 3 — explore & summarize. 'Sales' is a Table with a Revenue column. "
                     "Start by filtering (Ctrl+Shift+L) and reading the Status Bar; then write "
                     "each formula in column C. Commission uses the 'Reps' rate.")
        task_header(ex, 3, ["#", "Question", "Your formula"], [4, 74, 56])
        for i, (label, sol) in enumerate(STAGE3_PROMPTS, start=1):
            row = 3 + i
            ex.cell(row=row, column=1, value=i)
            ex.cell(row=row, column=2, value=label)
            if fill:
                cell = ex.cell(row=row, column=3, value=sol)
                if "rate" in label.lower():
                    cell.number_format = "0.0%"
                elif "Commission" in label:
                    cell.number_format = '"$"#,##0'
                elif "Order Value" in label or "revenue" in label.lower():
                    cell.number_format = '"$"#,##0'
        out = (SOLUTIONS_DIR if variant == "solution" else WORKING_DIR) / "module-3.xlsx"
        wb.save(out)


# ===================== Stage 4 · Pivot & rank ===============================

def build_stage4(rows: list[dict]) -> None:
    cleaned = _cleaned_rows(rows)

    # ---- working ----
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    write_sales_table(ws, cleaned, table_name="Sales", include_revenue=True)
    pad = wb.create_sheet("Pivot Scratchpad", 0)
    add_note(pad, "Stage 4 — build your PivotTables here (Insert > PivotTable). Add Slicers for "
                  "Region and Status. The last task is the sales-rep leaderboard.")
    task_list(pad, [
        "Pivot 1 — Revenue by Region (Rows = Region, Values = Sum of Revenue).",
        "Pivot 2 — Revenue by Region x Quarter (group OrderDate by Quarter into Columns).",
        "Pivot 3 — Revenue by Category, with a Status slicer.",
        "Pivot 4 — % of Grand Total: drag Revenue in twice, Show Values As > % of Grand Total.",
        "Leaderboard — Top 5 Sales Reps by Closed-Won revenue (Value Filter > Top 10 > set to 5), "
        "sorted largest first.",
    ])
    wb.save(WORKING_DIR / "module-4.xlsx")

    # ---- solution ----
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    write_sales_table(ws, cleaned, table_name="Sales", include_revenue=True)

    summary = wb.create_sheet("Summary", 0)
    add_note(summary, "Stage 4 solution. PivotTables/Slicers can't be authored by openpyxl, so "
                      "these SUMIFS tables mirror what your pivots should show.")
    regions = sorted({r["Region"] for r in cleaned})

    section(summary, 3, "Revenue by Region")
    summary["A4"] = "Region"; summary["B4"] = "Revenue"
    for c in ("A4", "B4"):
        summary[c].font = HEADER_FONT; summary[c].fill = HEADER_FILL
    for i, region in enumerate(regions, start=5):
        summary.cell(row=i, column=1, value=region)
        summary.cell(row=i, column=2, value=f'=SUMIFS(Sales[Revenue],Sales[Region],A{i})')
        summary.cell(row=i, column=2).number_format = '"$"#,##0'

    section(summary, 12, "Revenue by Category x Region")
    summary["A13"] = "Category"
    summary["A13"].font = HEADER_FONT; summary["A13"].fill = HEADER_FILL
    for j, region in enumerate(regions, start=2):
        c = summary.cell(row=13, column=j, value=region)
        c.font = HEADER_FONT; c.fill = HEADER_FILL
    for i, cat in enumerate(["Hardware", "Accessories", "Software", "Services"], start=14):
        summary.cell(row=i, column=1, value=cat)
        for j, region in enumerate(regions, start=2):
            col = get_column_letter(j)
            summary.cell(row=i, column=j,
                         value=f'=SUMIFS(Sales[Revenue],Sales[Category],$A{i},Sales[Region],{col}$13)')
            summary.cell(row=i, column=j).number_format = '"$"#,##0'

    # Sales rep leaderboard with data bars
    section(summary, 20, "Sales Rep Leaderboard  (Closed Won revenue, ranked, data bars)")
    summary["A21"] = "SalesRep"; summary["B21"] = "Closed Won Revenue"
    for c in ("A21", "B21"):
        summary[c].font = HEADER_FONT; summary[c].fill = HEADER_FILL
    cw_totals: dict[str, float] = {}
    for r in cleaned:
        if r["Status"] == "Closed Won":
            rev = r["UnitPrice"] * r["Quantity"] * (1 - r["Discount"])
            cw_totals[r["SalesRep"]] = cw_totals.get(r["SalesRep"], 0.0) + rev
    ranked = sorted(cw_totals.items(), key=lambda kv: -kv[1])
    for i, (rep, _t) in enumerate(ranked, start=22):
        summary.cell(row=i, column=1, value=rep)
        summary.cell(row=i, column=2,
                     value=f'=SUMIFS(Sales[Revenue],Sales[SalesRep],A{i},Sales[Status],"Closed Won")')
        summary.cell(row=i, column=2).number_format = '"$"#,##0'
    summary.conditional_formatting.add(
        f"B22:B{21 + len(ranked)}",
        DataBarRule(start_type="min", end_type="max", color="638EC6"))

    summary.column_dimensions["A"].width = 22
    for col in "BCDEFG":
        summary.column_dimensions[col].width = 16
    wb.save(SOLUTIONS_DIR / "module-4.xlsx")


# ===================== Stage 5 · Present ====================================
# One-page summary for the VP: KPI strip + one chart + conditional formatting.

def build_stage5(rows: list[dict]) -> None:
    cleaned = _cleaned_rows(rows)

    # ---- working ----
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    write_sales_table(ws, cleaned, table_name="Sales", include_revenue=True)
    pg = wb.create_sheet("One-Pager", 0)
    add_note(pg, "Stage 5 — build a one-page summary for the VP on this sheet. Keep it to one "
                 "screen, format money cleanly, and let conditional formatting do the flagging.")
    task_list(pg, [
        "KPI strip: Total Closed-Won revenue, # orders, Average Order Value, Win rate, Refund rate.",
        "One chart: a column chart of revenue by Region (sorted, titled, currency axis).",
        "Heat map: revenue by Category x Region with a colour scale (Conditional Formatting).",
        "Data bars on a rep-revenue column for an instant ranking.",
        "Flag rules: highlight duplicate Customer names; red-fill rows where Status = Refunded.",
        "Format money as $1.2M with a custom format; fit everything on one screen at 100%.",
    ])
    wb.save(WORKING_DIR / "module-5.xlsx")

    # ---- solution ----
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    write_sales_table(ws, cleaned, table_name="Sales", include_revenue=True)

    dash = wb.create_sheet("One-Pager", 0)
    add_note(dash, "Stage 5 solution — a one-page summary. KPI strip + a region chart + a "
                   "Category x Region heat map, all from live SUMIFS/COUNTIFS.")
    # KPI strip
    kpis = [
        ("Closed-Won Revenue", '=SUMIFS(Sales[Revenue],Sales[Status],"Closed Won")', '"$"#,##0'),
        ("Orders", '=COUNTA(Sales[OrderID])', "#,##0"),
        ("Avg Order Value",
         '=SUMIFS(Sales[Revenue],Sales[Status],"Closed Won")/COUNTIFS(Sales[Status],"Closed Won")', '"$"#,##0'),
        ("Win Rate", '=COUNTIFS(Sales[Status],"Closed Won")/COUNTA(Sales[OrderID])', "0.0%"),
        ("Refund Rate", '=COUNTIFS(Sales[Status],"Refunded")/COUNTA(Sales[OrderID])', "0.0%"),
    ]
    for i, (label, formula, fmt) in enumerate(kpis):
        col = 1 + i * 2
        dash.cell(row=3, column=col, value=label).font = Font(bold=True, color="1F4E78")
        cell = dash.cell(row=4, column=col, value=formula)
        cell.font = Font(bold=True, size=14)
        cell.number_format = fmt
        dash.column_dimensions[get_column_letter(col)].width = 20

    # Revenue by Region + chart
    section(dash, 7, "Revenue by Region")
    dash["A8"] = "Region"; dash["B8"] = "Revenue"
    for c in ("A8", "B8"):
        dash[c].font = HEADER_FONT; dash[c].fill = HEADER_FILL
    regions = sorted({r["Region"] for r in cleaned})
    for i, region in enumerate(regions, start=9):
        dash.cell(row=i, column=1, value=region)
        dash.cell(row=i, column=2, value=f'=SUMIFS(Sales[Revenue],Sales[Region],A{i})')
        dash.cell(row=i, column=2).number_format = '"$"#,##0'
    chart = BarChart(); chart.type = "col"; chart.title = "Revenue by Region"
    data = Reference(dash, min_col=2, min_row=8, max_row=8 + len(regions))
    cats = Reference(dash, min_col=1, min_row=9, max_row=8 + len(regions))
    chart.add_data(data, titles_from_data=True); chart.set_categories(cats)
    chart.height = 8; chart.width = 14
    dash.add_chart(chart, "D7")

    # Category x Region heat map
    section(dash, 17, "Revenue by Category x Region  (colour scale = heat map)")
    dash["A18"] = "Category"
    dash["A18"].font = HEADER_FONT; dash["A18"].fill = HEADER_FILL
    for j, region in enumerate(regions, start=2):
        c = dash.cell(row=18, column=j, value=region)
        c.font = HEADER_FONT; c.fill = HEADER_FILL
    for i, cat in enumerate(["Hardware", "Accessories", "Software", "Services"], start=19):
        dash.cell(row=i, column=1, value=cat)
        for j, region in enumerate(regions, start=2):
            col = get_column_letter(j)
            dash.cell(row=i, column=j,
                      value=f'=SUMIFS(Sales[Revenue],Sales[Category],$A{i},Sales[Region],{col}$18)')
            dash.cell(row=i, column=j).number_format = '"$"#,##0'
    end_col = get_column_letter(1 + len(regions))
    dash.conditional_formatting.add(
        f"B19:{end_col}22",
        ColorScaleRule(start_type="min", start_color="FFFFFF",
                       mid_type="percentile", mid_value=50, mid_color="FFEB84",
                       end_type="max", end_color="63BE7B"))
    dash.column_dimensions["A"].width = 16
    wb.save(SOLUTIONS_DIR / "module-5.xlsx")


# =============================== Capstone ===================================
# Fresh file: run the whole workflow and answer ~11 business questions.

CAPSTONE_QUESTIONS: list[tuple[str, str, str]] = [
    ("1. Total revenue from Closed Won orders only?",
     '=SUMIFS(Sales[Revenue],Sales[Status],"Closed Won")', '"$"#,##0'),
    ("2. How many orders are there in total?",
     '=COUNTA(Sales[OrderID])', "#,##0"),
    ("3. Refund rate = Refunded orders / all orders?",
     '=COUNTIFS(Sales[Status],"Refunded")/COUNTA(Sales[OrderID])', "0.0%"),
    ("4. Win rate = Closed Won orders / all orders?",
     '=COUNTIFS(Sales[Status],"Closed Won")/COUNTA(Sales[OrderID])', "0.0%"),
    ("5. Average Order Value on Closed Won orders?",
     '=SUMIFS(Sales[Revenue],Sales[Status],"Closed Won")/COUNTIFS(Sales[Status],"Closed Won")',
     '"$"#,##0'),
    ("6. Closed Won revenue for the West region?",
     '=SUMIFS(Sales[Revenue],Sales[Region],"West",Sales[Status],"Closed Won")', '"$"#,##0'),
    ("7. Closed Won revenue for the Hardware category?",
     '=SUMIFS(Sales[Revenue],Sales[Category],"Hardware",Sales[Status],"Closed Won")', '"$"#,##0'),
    ("8. Closed Won revenue booked by Anna Becker?",
     '=SUMIFS(Sales[Revenue],Sales[SalesRep],"Anna Becker",Sales[Status],"Closed Won")', '"$"#,##0'),
    ("9. Anna Becker's quota attainment? (her Closed Won revenue / her AnnualQuota from Reps)",
     '=SUMIFS(Sales[Revenue],Sales[SalesRep],"Anna Becker",Sales[Status],"Closed Won")'
     '/VLOOKUP("Anna Becker",Reps,4,FALSE)', "0.0%"),
    ("10. Commission earned by Anna Becker? (her Closed Won revenue x her CommissionRate from Reps)",
     '=SUMIFS(Sales[Revenue],Sales[SalesRep],"Anna Becker",Sales[Status],"Closed Won")'
     '*VLOOKUP("Anna Becker",Reps,5,FALSE)', '"$"#,##0'),
    ("11. Average discount given on Closed Won orders?",
     '=AVERAGEIFS(Sales[Discount],Sales[Status],"Closed Won")', "0.0%"),
]


def build_capstone(rows: list[dict], reps: list[dict]) -> None:
    cleaned = _cleaned_rows(rows)

    def base(wb):
        ws = wb.active
        ws.title = "Sales"
        write_sales_table(ws, cleaned, table_name="Sales", include_revenue=True)
        rs = wb.create_sheet("Reps")
        write_reps_table(rs, reps, table_name="Reps")

    # ---- working ----
    wb = Workbook()
    base(wb)
    q = wb.create_sheet("Questions", 0)
    add_note(q, "CAPSTONE — a fresh file. Run the whole workflow: it's already clean, so enrich "
                "with the 'Reps' table where needed and answer each question in column C. Use "
                "PivotTables, AutoFilter, SUMIFS/COUNTIFS, and VLOOKUP — whatever is fastest.")
    task_header(q, 3, ["#", "Business question", "Your answer"], [4, 78, 26])
    for i, (question, _sol, _fmt) in enumerate(CAPSTONE_QUESTIONS, start=1):
        q.cell(row=3 + i, column=1, value=i)
        q.cell(row=3 + i, column=2, value=question)
        q.cell(row=3 + i, column=3).fill = TASK_FILL
    wb.save(WORKING_DIR / "capstone.xlsx")

    # ---- solution ----
    wb = Workbook()
    base(wb)
    a = wb.create_sheet("Answers", 0)
    add_note(a, "CAPSTONE answer key. Each answer is a live formula so it stays correct if the "
                "data changes. Many also have a valid PivotTable route.")
    task_header(a, 3, ["#", "Business question", "Answer"], [4, 78, 26])
    for i, (question, sol, fmt) in enumerate(CAPSTONE_QUESTIONS, start=1):
        a.cell(row=3 + i, column=1, value=i)
        a.cell(row=3 + i, column=2, value=question)
        cell = a.cell(row=3 + i, column=3, value=sol)
        cell.number_format = fmt
    wb.save(SOLUTIONS_DIR / "capstone.xlsx")


# ------------------------------ entry ---------------------------------------

def main() -> None:
    WORKING_DIR.mkdir(parents=True, exist_ok=True)
    SOLUTIONS_DIR.mkdir(parents=True, exist_ok=True)

    rows = load_rows()
    reps = load_reps()
    print(f"Loaded {len(rows)} source rows, {len(reps)} reps")

    build_stage1(rows);          print("  stage 1 (module-1) done")
    build_stage2(rows, reps);    print("  stage 2 (module-2) done")
    build_stage3(rows, reps);    print("  stage 3 (module-3) done")
    build_stage4(rows);          print("  stage 4 (module-4) done")
    build_stage5(rows);          print("  stage 5 (module-5) done")
    build_capstone(rows, reps);  print("  capstone done")


if __name__ == "__main__":
    main()
