---
layout: lesson
title: "Stage 4 — Pivot & rank"
short_title: "Pivot & rank"
subtitle: "Reshape the whole dataset in seconds — and rank reps and brand managers."
module: 4
---

> **Where we are:** Receive & clean → Structure & enrich → Explore & summarize → **Pivot & rank** → Present.

PivotTables are the fastest way to slice your `Sales` table by area, category, brand, or customer — no formulas at all. Open `module-4.xlsx` and build on the `Pivot Scratchpad`.

## 1. Build your first pivot

Click any cell in `Sales` → **Insert → PivotTable → New Worksheet → OK**. Then:

- Drag **Area** to **Rows**
- Drag **SalesValue** to **Values**

Net sales by area. Done. (Because returns are negative, the pivot's sum is already *net*.)

## 2. The four drop zones

| Area | Use it for |
|------|------------|
| **Rows** | Categories down the side (Area, Category, Brand, Customer, SalesRep, BrandManager) |
| **Columns** | A second dimension across the top (often a time bucket) |
| **Values** | The numbers to summarize (SalesValue, SalesQuantity) |
| **Filters** | Page-level filters (InvoiceType, Category) |

Values defaults to **Sum** — click the field → **Value Field Settings** for Count, Average, etc.

## 3. Group dates, show shares

- Drag **Date** to Rows → right-click a date → **Group** → tick **Quarters** and **Years**. (Works because you fixed the text-dates in Stage 1.)
- Right-click a value → **Show Values As → % of Grand Total** to see each brand's share. Drag `SalesValue` into Values **twice** for AED and % side by side.

## 4. Slicers — clickable filters (and the dashboard trick)

**PivotTable Analyze → Insert Slicer** for `Area` and `InvoiceType`. The power move: one slicer can drive **several** pivots — right-click → **Report Connections** → tick the pivots. One click on "Deira" updates them all.

## 5. The leaderboards

Two your manager will ask for:

**Top reps by net sales** — pivot of `SalesRep → SalesValue`, then right-click a rep → **Filter → Top 10** → set to **5**, and **Sort → Largest to Smallest**.

**Net sales by brand manager** — pivot of `BrandManager → SalesValue` (the column you VLOOKUP'd in Stage 2). Instantly shows which manager's brands are driving the business.

The solution file's `Summary` sheet mirrors both with `SUMIFS` and adds **data bars** for a visual ranking.

## 6. Refreshing

Pivots don't update on their own. When `Sales` changes, right-click → **Refresh** (or **Data → Refresh All**). Because `Sales` is a Table, new rows are picked up automatically.

## Practice

On the scratchpad: net sales by Area; Category × Quarter; net sales by Customer with an `InvoiceType` slicer; % of Grand Total by Brand; the **Top-5 rep leaderboard**; and **net sales by Brand Manager**.

## Hand-off

You've got the numbers and the rankings. The last step is making them land for someone with ten seconds and a meeting to run.

## Further learning

**Official Microsoft docs**
- [Create a PivotTable to analyze worksheet data](https://support.microsoft.com/office/create-a-pivottable-to-analyze-worksheet-data-a9a84538-bfe9-40a9-a8e9-f99134456576)
- [Group or ungroup data in a PivotTable](https://support.microsoft.com/en-us/office/group-or-ungroup-data-in-a-pivottable-c9d1ddd0-6580-47d1-82bc-c84a5a340725)
- [Use slicers to filter data](https://support.microsoft.com/en-us/office/use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)

**Video:** [PivotTables — video tutorials](https://www.youtube.com/results?search_query=excel+pivot+table+tutorial)

Take the [Stage 4 quiz](../quizzes/module-4), then finish with [Stage 5 — Present & capstone](05-present-capstone).
