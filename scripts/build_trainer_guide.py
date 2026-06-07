"""Build the Trainer's Guide (facilitator handbook) as a polished PDF.

Reads the synthetic datasets, computes the capstone answer key, renders a styled
HTML handbook, and prints it to PDF with headless Chrome/Edge.

Output:
- ../docs/files/trainer-guide.html
- ../docs/files/excel-for-sales-trainer-guide.pdf   (served on GitHub Pages)
"""

from __future__ import annotations

import csv
import datetime
import html
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "docs" / "files" / "source"
OUT_HTML = ROOT / "docs" / "files" / "trainer-guide.html"
OUT_PDF = ROOT / "docs" / "files" / "excel-for-sales-trainer-guide.pdf"

TRAINER = "Oleks Tverdokhlieb"
REPO = "github.com/scripts-and-tables/Excel-101"
REPO_URL = "https://github.com/scripts-and-tables/Excel-101"
SITE_URL = "https://scripts-and-tables.github.io/Excel-101/"
VERSION = "Version 1.0"


# ----------------------------- data + answer key ----------------------------

def load(name):
    with (SRC / name).open(encoding="utf-8") as f:
        return list(csv.DictReader(f))


def compute():
    rows = load("sales_data.csv")
    reps = {r["SalesRep"]: int(r["AnnualQuota"]) for r in load("reps.csv")}
    brands = load("brands.csv")
    brand_mgr = {b["Brand"]: b["BrandManager"] for b in brands}

    # de-dup exactly like build_workbooks._cleaned_rows
    seen, c = set(), []
    for r in rows:
        cust = r["Customer"].strip()
        area = r["Area"].title()
        key = (r["Date"], r["InvoiceType"], r["CustomerCode"], r["BranchCode"],
               r["SalesRep"], r["Brand"], r["SalesQuantity"], r["SalesValue"])
        if key in seen:
            continue
        seen.add(key)
        r = dict(r); r["Customer"] = cust; r["Area"] = area
        r["v"] = float(r["SalesValue"])
        c.append(r)

    def s(pred):
        return sum(r["v"] for r in c if pred(r))

    gross = s(lambda r: r["InvoiceType"] == "Sales")
    returns = s(lambda r: r["InvoiceType"] == "Return")
    net = s(lambda r: True)
    sales_rows = [r for r in c if r["InvoiceType"] == "Sales"]
    ms_net = s(lambda r: r["SalesRep"] == "Mohammed Saleh")
    ms_quota = reps["Mohammed Saleh"]
    return {
        "rows": len(c),
        "customers": len({r["Customer"] for r in c}),
        "branches": len({r["BranchCode"] for r in c}),
        "areas": len({r["Area"] for r in c}),
        "reps": len(reps),
        "brands": len(brands),
        "brand_mgrs": len({b["BrandManager"] for b in brands}),
        "date_min": min(r["Date"] for r in c),
        "date_max": max(r["Date"] for r in c),
        "net": net, "gross": gross, "returns": returns,
        "return_rate": (-returns / gross) if gross else 0,
        "n_invoices": len(c),
        "avg_sale": (sum(r["v"] for r in sales_rows) / len(sales_rows)) if sales_rows else 0,
        "deira": s(lambda r: r["Area"] == "Deira"),
        "food": s(lambda r: r["Category"] == "Food"),
        "carrefour": s(lambda r: r["Customer"] == "Carrefour"),
        "imran": s(lambda r: brand_mgr.get(r["Brand"]) == "Imran Sheikh"),
        "ms_net": ms_net, "ms_quota": ms_quota,
        "ms_attain": (ms_net / ms_quota) if ms_quota else 0,
    }


def aed(x):
    return f"AED&nbsp;{round(x):,}"


def pct(x):
    return f"{x * 100:.1f}%"


# ----------------------------- content -------------------------------------

MODULES = [
    {
        "n": 1, "title": "Receive &amp; clean", "time": "50 min",
        "file": "module-1.xlsx",
        "obj": "Turn a raw supermarket sales export into a clean, trustworthy Excel Table.",
        "teach": [
            "Make it a Table (<b>Ctrl+T</b>), name it <code>Sales</code> — structured references, auto-expand, filters.",
            "<b>TRIM</b> stray spaces, <b>CLEAN</b> hidden line-breaks/tabs in Customer names.",
            "<b>Find &amp; Replace</b> (Ctrl+H) to fix CAPS areas in bulk (DEIRA &rarr; Deira).",
            "<b>VALUE</b> for numbers stored as text, <b>DATEVALUE</b> for text dates (left-aligned = text).",
            "Split/combine codes: <b>LEFT / MID / RIGHT</b>, <b>SUBSTITUTE</b>, <b>&amp;</b> to join.",
            "<b>Paste Special &rarr; Values</b> to lock helper results; Remove Duplicates; sort; freeze header.",
            "Number formats: AED currency &amp; the <code>\"AED\" #,##0</code> custom format.",
        ],
        "watch": [
            "Returns carry <b>negative</b> quantity &amp; value — point this out early; it underpins every KPI later.",
            "The most common 'my SUM is wrong' is text-stored numbers — show the left-align tell.",
            "Data &gt; Remove Duplicates keeps the first occurrence; untick OrderNumber when de-duping.",
        ],
        "ex": "RawData (messy) + a Cleanup Practice sheet (one cell per function) + 10 exercises. Solution = clean Sales Table.",
    },
    {
        "n": 2, "title": "Structure &amp; enrich", "time": "45 min",
        "file": "module-2.xlsx",
        "obj": "Add the context the export lacks — sales manager &amp; quota (Reps) and brand manager (Brands) — with VLOOKUP.",
        "teach": [
            "<b>VLOOKUP</b> anatomy: <code>=VLOOKUP([@SalesRep], Reps, 2, FALSE)</code> &rarr; Manager; <code>3</code> &rarr; Quota.",
            "Second lookup: <code>=VLOOKUP([@Brand], Brands, 3, FALSE)</code> &rarr; BrandManager.",
            "Always <b>FALSE</b> (exact). Demonstrate how TRUE silently returns the wrong row.",
            "The <b>six failure modes</b> (forgot FALSE, can't look left, column-index shift, text-vs-number, spaces, genuine #N/A).",
            "<b>IFERROR</b> to turn #N/A into \"unknown\". XLOOKUP shown as the modern note only.",
        ],
        "watch": [
            "Learners conflate lookup value vs first column — stress VLOOKUP matches the table's <i>first</i> column.",
            "If a lookup #N/As on identical-looking text, it's almost always spaces or text-vs-number (ties back to Stage 1).",
            "Keep XLOOKUP brief — many learners are on Excel 2016/2019 where it doesn't exist.",
        ],
        "ex": "Sales + Reps + Brands sheets, 7 exercises. Solution = enriched Sales (SalesManager/BrandManager/Quota) + a Failure Modes sheet.",
    },
    {
        "n": 3, "title": "Explore &amp; summarize", "time": "60 min",
        "file": "module-3.xlsx",
        "obj": "Answer everyday questions: filter first, then the conditional trio and the sales KPIs.",
        "teach": [
            "Explore with <b>AutoFilter</b> (Ctrl+Shift+L) and the <b>Status Bar</b> (sum/count/avg, no formula).",
            "<b>SUBTOTAL(9/103/101, …)</b> totals only visible rows; Data &gt; Subtotal for group totals (range only).",
            "The trio: <b>SUMIFS / COUNTIFS / AVERAGEIFS</b> — sum_range first, then criteria pairs.",
            "Net sales = <code>SUM(SalesValue)</code> (returns net out automatically); gross = SUMIFS on InvoiceType=Sales.",
            "KPIs: <b>return rate</b>, <b>average sale</b>, <b>sales by brand manager</b> (enriched column), <b>quota attainment</b> (SUMIFS ÷ VLOOKUP).",
            "Mechanics: double-click fill handle; F4 absolute/relative.",
        ],
        "watch": [
            "Return rate: returns are negative, so negate before dividing — <code>=-SUMIFS(…Return)/SUMIFS(…Sales)</code>.",
            "Data &gt; Subtotal is greyed out on a Table — convert to range first (common 'why can't I' moment).",
            "Encourage criteria-in-cells so a formula becomes a reusable mini-report.",
        ],
        "ex": "Enriched Sales + Reps + Brands, 11 prompts (SUM/AVERAGE/COUNT, IF, the trio, KPIs, commission-free quota).",
    },
    {
        "n": 4, "title": "Pivot &amp; rank", "time": "40 min",
        "file": "module-4.xlsx",
        "obj": "Reshape the whole dataset in seconds and produce the leaderboards.",
        "teach": [
            "Build a pivot; the four drop zones; switch Sum &rarr; Count/Average via Value Field Settings.",
            "Group <b>Date</b> by Quarter/Year; <b>Show Values As &rarr; % of Grand Total</b>.",
            "<b>Slicers</b> for Area/InvoiceType; <b>Report Connections</b> to drive several pivots (mini-dashboard).",
            "Leaderboards: <b>Top-5 reps</b> by net sales (Value Filter &rarr; Top 10 &rarr; 5) and <b>net sales by Brand Manager</b>.",
            "Refresh: pivots cache — right-click &rarr; Refresh / Data &gt; Refresh All.",
        ],
        "watch": [
            "Date grouping needs real dates — if grouping fails, a date column is still text (back to Stage 1).",
            "openpyxl can't author pivots, so the solution mirrors them with SUMIFS + data bars — tell learners that.",
        ],
        "ex": "Enriched Sales + scratchpad (4 pivots + 2 leaderboards). Solution = SUMIFS mirror with a data-bar leaderboard.",
    },
    {
        "n": 5, "title": "Present &amp; capstone", "time": "70 min",
        "file": "module-5.xlsx + capstone.xlsx",
        "obj": "Turn the analysis into a one-page summary, then have learners run the whole workflow on a fresh file.",
        "teach": [
            "KPI strip (net sales, gross, return rate, avg sale, # invoices) — most important number top-left.",
            "One clean column chart (sorted, titled, AED axis); a Category × Area heat map (color scale).",
            "Conditional formatting: data bars, icon sets, duplicate values, and a formula rule to flag Return rows red.",
            "Hand out the capstone: a fresh file + the 11 questions; let them work, then walk the answer key.",
        ],
        "watch": [
            "Keep presentation ruthless — one screen, one colour per area, AED 1.2M not 1234567.89.",
            "In the capstone, push live formulas over typed numbers so answers survive a data refresh.",
        ],
        "ex": "One-Pager build + capstone.xlsx (Questions / Answers). Answer key with expected values is in this guide.",
    },
]

AGENDA = [
    ("Intro &amp; setup", "10 min"),
    ("M1 · Receive &amp; clean", "50 min"),
    ("M2 · Structure &amp; enrich", "45 min"),
    ("☕ Break", "10 min"),
    ("M3 · Explore &amp; summarize", "60 min"),
    ("M4 · Pivot &amp; rank", "40 min"),
    ("☕ Break", "10 min"),
    ("M5 · Present", "50 min"),
    ("🎯 Capstone", "40 min"),
    ("Wrap-up &amp; Q&amp;A", "10 min"),
]

SHORTCUTS = [
    ("Ctrl + T", "Create a Table"), ("Ctrl + Shift + L", "Toggle AutoFilter"),
    ("Ctrl + Alt + V", "Paste Special"), ("Ctrl + H / Ctrl + F", "Replace / Find"),
    ("Ctrl + D", "Fill down"), ("Ctrl + Shift + ↓", "Select to bottom of column"),
    ("Ctrl + 1", "Format Cells"), ("Ctrl + Shift + 1", "Number format (thousands)"),
    ("Alt + =", "AutoSum"), ("Alt + N + V", "Insert PivotTable"),
    ("F2", "Edit active cell"), ("F4", "Toggle absolute / relative reference"),
]

DOC_LINKS = {
    "M1 — cleaning & text": [
        ("Create and format an Excel table", "https://support.microsoft.com/office/create-and-format-tables-e81aa349-b006-4f8a-9806-5af9df0ac664"),
        ("Top ten ways to clean your data", "https://support.microsoft.com/en-us/office/top-ten-ways-to-clean-your-data-2844b620-677c-47a7-ac3e-c2e157d1db19"),
        ("TRIM", "https://support.microsoft.com/office/410388fa-c5df-49c6-b16c-9e5630b479f9"),
        ("CLEAN", "https://support.microsoft.com/office/26f3d7c5-475f-4a9c-90e5-4b8ba987ba41"),
        ("VALUE", "https://support.microsoft.com/office/257d0108-07dc-437d-ae1c-bc2d3953d8c2"),
        ("DATEVALUE", "https://support.microsoft.com/office/df8b07d4-7761-4a93-bc33-b7471bbff252"),
        ("LEFT", "https://support.microsoft.com/office/9203d2d2-7960-479b-84c6-1ea52b99640c"),
        ("MID", "https://support.microsoft.com/office/d5f9e25c-d7d6-472e-b568-4ecb12433028"),
        ("RIGHT", "https://support.microsoft.com/office/240267ee-9afa-4639-a02b-f19e1786cf2f"),
        ("SUBSTITUTE", "https://support.microsoft.com/office/6434944e-a904-4336-a9b0-1e58df3bc332"),
        ("Combine text (&amp; / CONCAT)", "https://support.microsoft.com/en-us/office/combine-text-from-two-or-more-cells-into-one-cell-81ba0946-ce78-42ed-b3c3-21340eb164a6"),
        ("Find or replace text and numbers", "https://support.microsoft.com/en-us/office/find-or-replace-text-and-numbers-on-a-worksheet-0e304ca5-ecef-4808-b90f-fdb42f892e90"),
        ("Find and remove duplicates", "https://support.microsoft.com/en-us/office/find-and-remove-duplicates-00e35bea-b46a-4d5d-b28e-66a552dc138d"),
        ("Sort data in a range or table", "https://support.microsoft.com/en-us/office/sort-data-in-a-range-or-table-in-excel-62d0b95d-2a90-4610-a6ae-2e545c4a4654"),
        ("Freeze panes", "https://support.microsoft.com/en-us/office/freeze-panes-to-lock-rows-and-columns-dab2ffc9-020d-4026-8121-67dd25f2508f"),
        ("Move or copy cells (Paste Special)", "https://support.microsoft.com/en-us/office/move-or-copy-cells-rows-and-columns-3ebbcafd-8566-42d8-8023-a2ec62746cfc"),
        ("Available number formats", "https://support.microsoft.com/en-us/office/available-number-formats-in-excel-0afe8f52-97db-41f1-b972-4b46e9f1e8d2"),
    ],
    "M2 — lookups": [
        ("VLOOKUP", "https://support.microsoft.com/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1"),
        ("VLOOKUP troubleshooting (failure modes)", "https://support.microsoft.com/en-us/office/quick-reference-card-vlookup-troubleshooting-tips-6fe7fe1b-709b-4958-adfb-9f2a409dcf38"),
        ("IFERROR", "https://support.microsoft.com/en-us/office/iferror-function-c526fd07-caeb-47b8-8bb6-63f3e417f611"),
        ("XLOOKUP", "https://support.microsoft.com/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929"),
        ("Look up values with VLOOKUP, INDEX or MATCH", "https://support.microsoft.com/en-us/office/look-up-values-with-vlookup-index-or-match-68297403-7c3c-4150-9e3c-4d348188976b"),
        ("MATCH", "https://support.microsoft.com/office/e8dffd45-c762-47d6-bf89-533f4a37673a"),
    ],
    "M3 — summarizing": [
        ("SUM", "https://support.microsoft.com/office/043e1c7d-7726-4e80-8f32-07b23e057f89"),
        ("AVERAGE", "https://support.microsoft.com/office/047bac88-d466-426c-a32b-8f33eb960cf6"),
        ("COUNT", "https://support.microsoft.com/office/a59cd7fc-b623-4d93-87a4-d23bf411294c"),
        ("COUNTA", "https://support.microsoft.com/office/7dc98875-d5c1-46f1-9a82-53f3219e2509"),
        ("MAX", "https://support.microsoft.com/office/e0012414-9ac8-4b34-9a47-73e662c08098"),
        ("IF", "https://support.microsoft.com/office/69aed7c9-4e8a-4755-a9bc-aa8bbff73be2"),
        ("SUMIFS", "https://support.microsoft.com/office/c9e748f5-7ea7-455d-9406-611cebce642b"),
        ("COUNTIFS", "https://support.microsoft.com/office/dda3dc6e-f74e-4aee-88bc-aa8c2a866842"),
        ("AVERAGEIFS", "https://support.microsoft.com/office/48910c45-1fc0-4389-a028-f7c5c3001690"),
        ("SUBTOTAL", "https://support.microsoft.com/office/7b027003-f060-4ade-9040-e478765b9939"),
        ("Filter data in a range or table", "https://support.microsoft.com/en-us/office/filter-data-in-a-range-or-table-in-excel-01832226-31b5-4568-8806-38c37dcc180e"),
        ("Insert subtotals (Data &gt; Subtotal)", "https://support.microsoft.com/en-us/office/insert-subtotals-in-a-list-of-data-in-a-worksheet-7881d256-b4fa-4f81-b71e-b0a3d4a52b3a"),
        ("Use AutoSum", "https://support.microsoft.com/en-us/office/use-autosum-to-sum-numbers-in-excel-543941e7-e783-44ef-8317-7d1bb85fe706"),
        ("Relative / absolute / mixed references", "https://support.microsoft.com/en-us/office/switch-between-relative-absolute-and-mixed-references-dfec08cd-ae65-4f56-839e-5f0d8d0baca9"),
    ],
    "M4 — pivot &amp; rank": [
        ("Create a PivotTable", "https://support.microsoft.com/office/create-a-pivottable-to-analyze-worksheet-data-a9a84538-bfe9-40a9-a8e9-f99134456576"),
        ("Group or ungroup data in a PivotTable", "https://support.microsoft.com/en-us/office/group-or-ungroup-data-in-a-pivottable-c9d1ddd0-6580-47d1-82bc-c84a5a340725"),
        ("Show Values As", "https://support.microsoft.com/en-us/office/show-different-calculations-in-pivottable-value-fields-014d2777-baaf-480b-a32b-98431f48bfec"),
        ("Use slicers to filter data", "https://support.microsoft.com/en-us/office/use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d"),
        ("Refresh PivotTable data", "https://support.microsoft.com/en-us/office/refresh-pivottable-data-6d24cece-a038-468a-8176-8b6568ca9be2"),
    ],
    "M5 — present": [
        ("Add, change, or clear conditional formats", "https://support.microsoft.com/office/add-change-or-clear-conditional-formats-fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f"),
        ("Create a chart from start to finish", "https://support.microsoft.com/en-us/office/create-a-chart-from-start-to-finish-0baf399e-dd61-4e18-8a73-b3fd5d5680c2"),
        ("Available number formats", "https://support.microsoft.com/en-us/office/available-number-formats-in-excel-0afe8f52-97db-41f1-b972-4b46e9f1e8d2"),
        ("Keyboard shortcuts in Excel", "https://support.microsoft.com/en-us/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f"),
    ],
}

CAPSTONE = [
    ("1", "Net sales (sum of SalesValue)", lambda d: aed(d["net"])),
    ("2", "Gross sales (Sales invoices only)", lambda d: aed(d["gross"])),
    ("3", "Total returns value", lambda d: aed(d["returns"])),
    ("4", "Return rate (returns ÷ gross)", lambda d: pct(d["return_rate"])),
    ("5", "Number of invoice lines", lambda d: f"{d['n_invoices']:,}"),
    ("6", "Average sale value (Sales lines only)", lambda d: aed(d["avg_sale"])),
    ("7", "Net sales in the Deira area", lambda d: aed(d["deira"])),
    ("8", "Net sales for the Food category", lambda d: aed(d["food"])),
    ("9", "Net sales for customer Carrefour", lambda d: aed(d["carrefour"])),
    ("10", "Net sales for brand manager Imran Sheikh", lambda d: aed(d["imran"])),
    ("11", "Mohammed Saleh — quota attainment", lambda d: pct(d["ms_attain"])),
]


# ----------------------------- HTML ----------------------------------------

def build_html(d):
    today = datetime.date.today().strftime("%-d %B %Y") if sys.platform != "win32" \
        else datetime.date.today().strftime("%d %B %Y")

    def li(items):
        return "".join(f"<li>{x}</li>" for x in items)

    agenda_rows = "".join(
        f"<tr><td>{b}</td><td class='r'>{t}</td></tr>" for b, t in AGENDA)
    shortcut_rows = "".join(
        f"<tr><td><kbd>{html.escape(k)}</kbd></td><td>{v}</td></tr>" for k, v in SHORTCUTS)
    capstone_rows = "".join(
        f"<tr><td class='r'>{n}</td><td>{q}</td><td class='r mono'>{fn(d)}</td></tr>"
        for n, q, fn in CAPSTONE)

    modules_html = ""
    for m in MODULES:
        modules_html += f"""
        <section class="module">
          <h3><span class="mnum">M{m['n']}</span> {m['title']}
              <span class="mtime">{m['time']}</span></h3>
          <p class="obj"><b>Objective.</b> {m['obj']}</p>
          <p class="files">Files: <code>{m['file']}</code></p>
          <div class="cols">
            <div>
              <h4>Teach</h4><ul>{li(m['teach'])}</ul>
            </div>
            <div>
              <h4>Watch for</h4><ul class="warn">{li(m['watch'])}</ul>
              <h4>Exercises</h4><p>{m['ex']}</p>
            </div>
          </div>
        </section>"""

    doc_html = ""
    for grp, links in DOC_LINKS.items():
        items = "".join(f"<li><a href='{u}'>{name}</a></li>" for name, u in links)
        doc_html += f"<h4>{grp}</h4><ul class='links'>{items}</ul>"

    return f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<title>Excel for Sales — Trainer's Guide</title>
<style>
  @page {{ size: A4; margin: 16mm 15mm 18mm 15mm; }}
  * {{ box-sizing: border-box; }}
  body {{ font-family: -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    color: #1f2a37; font-size: 10.5pt; line-height: 1.5; margin: 0; }}
  a {{ color: #047857; text-decoration: none; }}
  code, .mono, kbd {{ font-family: "Consolas", "SF Mono", Menlo, monospace; }}
  code {{ background: #ecfdf5; color: #065f46; padding: 0 3px; border-radius: 3px; font-size: 9.2pt; }}
  kbd {{ background: #f3f4f6; border: 1px solid #d8e0dc; border-radius: 4px; padding: 1px 6px; font-size: 9pt; }}

  /* running footer on every page */
  .footer {{ position: fixed; bottom: -12mm; left: 0; right: 0; font-size: 8pt;
    color: #6b7b73; border-top: 1px solid #e2e8e4; padding-top: 3px;
    display: flex; justify-content: space-between; }}

  /* cover */
  .cover {{ height: 247mm; display: flex; flex-direction: column; justify-content: center;
    page-break-after: always; }}
  .cover .bar {{ width: 64px; height: 6px; background: #047857; margin-bottom: 26px; }}
  .cover .eyebrow {{ color: #047857; font-weight: 700; letter-spacing: .14em;
    text-transform: uppercase; font-size: 11pt; }}
  .cover h1 {{ font-size: 40pt; margin: 6px 0 4px; letter-spacing: -1px; line-height: 1.05; }}
  .cover h2 {{ font-size: 17pt; font-weight: 600; color: #334155; margin: 0 0 28px; }}
  .cover .meta {{ font-size: 10.5pt; color: #475569; line-height: 1.9; }}
  .cover .meta b {{ color: #1f2a37; }}
  .cover .tag {{ margin-top: 30px; padding: 14px 18px; background: #ecfdf5;
    border-left: 4px solid #047857; border-radius: 6px; font-size: 9.5pt; color: #065f46; }}

  h2.sec {{ font-size: 17pt; color: #065f46; margin: 0 0 4px; padding-bottom: 6px;
    border-bottom: 2px solid #047857; page-break-after: avoid; }}
  .lead {{ color: #475569; margin: 0 0 14px; }}
  section.block {{ page-break-before: always; }}
  h3 {{ font-size: 12.5pt; margin: 16px 0 6px; page-break-after: avoid; }}
  h4 {{ font-size: 10pt; color: #047857; text-transform: uppercase; letter-spacing: .04em;
    margin: 12px 0 4px; }}
  ul {{ margin: 4px 0 8px; padding-left: 18px; }}
  li {{ margin-bottom: 3px; }}
  ul.warn li {{ }}
  table {{ width: 100%; border-collapse: collapse; margin: 8px 0 14px; font-size: 9.6pt; }}
  th, td {{ text-align: left; padding: 6px 9px; border-bottom: 1px solid #e2e8e4; vertical-align: top; }}
  thead th {{ background: #ecfdf5; color: #065f46; font-size: 8.6pt; text-transform: uppercase;
    letter-spacing: .03em; }}
  td.r, th.r {{ text-align: right; }}
  .grid2 {{ display: flex; gap: 22px; }}
  .grid2 > div {{ flex: 1; }}

  .module {{ border: 1px solid #e2e8e4; border-radius: 8px; padding: 12px 16px; margin: 0 0 14px;
    page-break-inside: avoid; }}
  .module h3 {{ margin-top: 0; display: flex; align-items: center; gap: 10px; }}
  .mnum {{ background: #047857; color: #fff; border-radius: 6px; padding: 2px 9px; font-size: 10pt; }}
  .mtime {{ margin-left: auto; color: #6b7b73; font-size: 9.5pt; font-weight: 400; }}
  .module .obj {{ margin: 2px 0 4px; }}
  .module .files {{ margin: 0 0 6px; color: #6b7b73; font-size: 9pt; }}
  .cols {{ display: flex; gap: 22px; }} .cols > div {{ flex: 1; }}
  ul.links {{ columns: 2; column-gap: 22px; }} ul.links li {{ break-inside: avoid; }}
  .kpibox {{ background: #f8fafb; border: 1px solid #e2e8e4; border-radius: 8px; padding: 10px 14px; }}
</style></head><body>

<div class="footer">
  <span>Excel for Sales · Trainer's Guide · {TRAINER}</span>
  <span>{REPO}</span>
</div>

<!-- COVER -->
<div class="cover">
  <div class="bar"></div>
  <div class="eyebrow">Trainer's Guide · Facilitator Handbook</div>
  <h1>Excel for Sales</h1>
  <h2>From a raw Dubai sales export to answers for the VP — the whole workflow.</h2>
  <div class="meta">
    <div><b>Prepared by</b> &nbsp; {TRAINER}</div>
    <div><b>Course repo</b> &nbsp; <a href="{REPO_URL}">{REPO}</a></div>
    <div><b>Live course</b> &nbsp; <a href="{SITE_URL}">{SITE_URL}</a></div>
    <div><b>Format</b> &nbsp; Intermediate · ≈ 5 hours · instructor-led, practice-heavy</div>
    <div><b>{VERSION}</b> &nbsp; · &nbsp; {today}</div>
  </div>
  <div class="tag"><b>100% synthetic data.</b> A UAE FMCG setting (fictional Food &amp; HPC brands
    sold to Dubai supermarkets). Customers are real chain names used only as labels; everything
    else is generated. Nothing real or confidential — and the brands are unrelated to any Transmed portfolio.</div>
</div>

<!-- OVERVIEW -->
<section class="block">
  <h2 class="sec">About this guide</h2>
  <p class="lead">This handbook is for <b>trainers delivering the course</b> — the run-sheet, talking
  points, pitfalls, and the capstone answer key. Learner-facing lessons, downloadable workbooks and
  quizzes live on the course site; this is the back-of-house companion.</p>

  <h3>Audience &amp; scope</h3>
  <p>Intermediate sales people (account managers, reps, sales ops) who <b>receive</b> sales exports and
  must analyse them. We teach the ~80% they use daily and deliberately exclude analyst tooling
  (dynamic arrays, Power Query, Power Pivot, macros). VLOOKUP is taught as primary; XLOOKUP is a note.
  Works in Excel 2016+.</p>

  <div class="grid2">
    <div>
      <h3>Timed agenda</h3>
      <table><tbody>{agenda_rows}</tbody></table>
      <p style="font-size:9pt;color:#6b7b73">Two short breaks; capstone is the assessment. Stretches to ~5 hrs with fuller Q&amp;A.</p>
    </div>
    <div>
      <h3>Learning outcomes</h3>
      <ul>
        <li>Clean a raw export into a trustworthy Table (TRIM/CLEAN, VALUE/DATEVALUE, text functions, dedupe).</li>
        <li>Enrich with VLOOKUP — rep&rarr;manager/quota and brand&rarr;brand manager — and dodge its six failures.</li>
        <li>Summarise with SUMIFS/COUNTIFS/AVERAGEIFS; compute net sales, return rate, AOV, quota attainment.</li>
        <li>Reshape and rank with PivotTables, slicers, leaderboards.</li>
        <li>Present a one-pager; run the whole workflow in a capstone.</li>
      </ul>
    </div>
  </div>
</section>

<!-- DATASET -->
<section class="block">
  <h2 class="sec">The dataset</h2>
  <p class="lead">One synthetic export flows through every module. Two reference tables drive the lookups.</p>
  <div class="grid2">
    <div>
      <h4>sales_data.csv — order lines</h4>
      <p style="font-size:9.4pt">OrderNumber · Date · InvoiceType (Sales/Return) · CustomerCode · Customer ·
      BranchCode · Branch · Area · SalesRep · Brand · Category · SalesQuantity · SalesValue.
      <b>Returns carry negative</b> quantity &amp; value.</p>
      <h4>reps.csv</h4><p style="font-size:9.4pt">SalesRep · Manager · AnnualQuota</p>
      <h4>brands.csv</h4><p style="font-size:9.4pt">Brand · Category · BrandManager</p>
    </div>
    <div class="kpibox">
      <h4>At a glance (computed)</h4>
      <table><tbody>
        <tr><td>Order lines</td><td class="r">{d['rows']:,}</td></tr>
        <tr><td>Period</td><td class="r">{d['date_min']} → {d['date_max']}</td></tr>
        <tr><td>Customers · branches · areas</td><td class="r">{d['customers']} · {d['branches']} · {d['areas']}</td></tr>
        <tr><td>Reps · brands · brand managers</td><td class="r">{d['reps']} · {d['brands']} · {d['brand_mgrs']}</td></tr>
        <tr><td>Net sales</td><td class="r">{aed(d['net'])}</td></tr>
        <tr><td>Return rate</td><td class="r">{pct(d['return_rate'])}</td></tr>
        <tr><td>Average sale value</td><td class="r">{aed(d['avg_sale'])}</td></tr>
      </tbody></table>
      <p style="font-size:8.5pt;color:#6b7b73;margin:4px 0 0">Example entities used in exercises:
      rep <b>Mohammed Saleh</b>, brand manager <b>Imran Sheikh</b>, area <b>Deira</b>,
      category <b>Food</b>, brand <b>Crunchio</b>, customer <b>Carrefour</b>.</p>
    </div>
  </div>
</section>

<!-- MODULE NOTES -->
<section class="block">
  <h2 class="sec">Module-by-module facilitator notes</h2>
  {modules_html}
</section>

<!-- CAPSTONE -->
<section class="block">
  <h2 class="sec">Capstone — answer key</h2>
  <p class="lead"><b>Brief.</b> "You're the new analyst on the Dubai account. Before Monday's meeting,
  answer the VP's questions, then drop the headline numbers into a one-page summary." Learners use
  PivotTables, AutoFilter, SUMIFS/COUNTIFS and VLOOKUP. Expected values below (your learners' numbers
  should match — the data is deterministic).</p>
  <table>
    <thead><tr><th class="r">#</th><th>Question</th><th class="r">Expected answer</th></tr></thead>
    <tbody>{capstone_rows}</tbody>
  </table>
  <p style="font-size:9pt;color:#6b7b73">Each answer also has a live-formula route in
  <code>capstone.xlsx</code> (Answers sheet) — prefer formulas over typed numbers so they survive a refresh.</p>
</section>

<!-- DELIVERY TIPS -->
<section class="block">
  <h2 class="sec">Delivery tips</h2>
  <ul>
    <li><b>Demo, then they do.</b> Each module: show the move once on the shared screen, then give them the exercise sheet and circulate.</li>
    <li><b>One running file.</b> Reinforce that it's the same dataset travelling the workflow — that's the spine that makes it stick.</li>
    <li><b>Returns are negative</b> is the single most important concept for the KPIs — anchor it in M1 and keep referring back.</li>
    <li><b>Virtual delivery:</b> share the working file in advance; use breakout time for exercises; reconvene to compare with the solution file.</li>
    <li><b>Pacing:</b> M3 is the longest and highest-value — protect its time; M4/M5 can compress if you're behind.</li>
    <li><b>Quizzes</b> (6 per module, in-browser, 70% to pass) are optional self-checks; the <b>capstone is the real assessment</b>.</li>
  </ul>
</section>

<!-- APPENDICES -->
<section class="block">
  <h2 class="sec">Appendix A · Keyboard shortcuts</h2>
  <table style="width:60%"><tbody>{shortcut_rows}</tbody></table>

  <h2 class="sec" style="margin-top:22px">Appendix B · Official Microsoft documentation</h2>
  <p class="lead">Every function/feature the course covers, linked to Microsoft's own docs.</p>
  {doc_html}

  <h2 class="sec" style="margin-top:22px">Appendix C · Downloads</h2>
  <p>All files are served from the course site under <code>/files/</code>:</p>
  <ul>
    <li>Working &amp; solution workbooks: <code>module-1.xlsx</code> … <code>module-5.xlsx</code>, <code>capstone.xlsx</code></li>
    <li>Datasets: <code>sales_data.csv</code>, <code>reps.csv</code>, <code>brands.csv</code></li>
    <li>Downloads page: <a href="{SITE_URL}downloads/">{SITE_URL}downloads/</a></li>
  </ul>
</section>

<!-- CREDITS -->
<section class="block">
  <h2 class="sec">Credits</h2>
  <div class="kpibox" style="max-width:120mm">
    <p style="margin:0 0 6px"><b>Author &amp; facilitator</b><br>{TRAINER}</p>
    <p style="margin:0 0 6px"><b>Course repository</b><br><a href="{REPO_URL}">{REPO_URL}</a></p>
    <p style="margin:0 0 6px"><b>Live course</b><br><a href="{SITE_URL}">{SITE_URL}</a></p>
    <p style="margin:0 0 6px"><b>Data</b><br>100% synthetic, generated from a fixed seed. Customers are
    real Dubai supermarket chain names used only as labels; brands are fictional and unrelated to any
    real company.</p>
    <p style="margin:0"><b>Licence</b><br>Course content MIT licensed. {VERSION} · {today}</p>
  </div>
</section>

</body></html>"""


def render_pdf():
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    browser = next((c for c in candidates if Path(c).exists()), None)
    if not browser:
        print("No Chrome/Edge found; HTML written, PDF skipped.")
        return False
    url = "file:///" + str(OUT_HTML).replace("\\", "/").replace(" ", "%20")
    args = [browser, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
            f"--print-to-pdf={OUT_PDF}", "--print-to-pdf-no-header", url]
    subprocess.run(args, check=True, timeout=180,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return OUT_PDF.exists()


def main():
    d = compute()
    OUT_HTML.write_text(build_html(d), encoding="utf-8")
    print(f"Wrote {OUT_HTML}")
    if render_pdf():
        kb = OUT_PDF.stat().st_size // 1024
        print(f"Wrote {OUT_PDF} ({kb} KB)")
    else:
        print("PDF not produced.")


if __name__ == "__main__":
    main()
