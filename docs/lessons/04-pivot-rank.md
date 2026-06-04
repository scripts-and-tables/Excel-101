---
layout: lesson
title: "Stage 4 — Pivot & rank"
short_title: "Pivot & rank"
subtitle: "Reshape and rank the whole dataset in seconds — and build the rep leaderboard."
module: 4
---

> **Where we are:** Receive & clean → Structure & enrich → Explore & summarize → **Pivot & rank** → Present.

PivotTables are the fastest way to slice your `Sales` table by region, period, rep, or product — no formulas at all. Drag a few fields and you've answered "what's our revenue by X?" in under a minute. Open `module-4.xlsx` and build on the `Pivot Scratchpad`.

## 1. Build your first pivot

Click any cell in `Sales` → **Insert → PivotTable → New Worksheet → OK**. Then:

- Drag **Region** to **Rows**
- Drag **Revenue** to **Values**

Revenue by region. Done.

## 2. The four drop zones

| Area | Use it for |
|------|------------|
| **Rows** | Categories down the side (Region, SalesRep, Product) |
| **Columns** | A second dimension across the top (often a time bucket) |
| **Values** | The numbers to summarize (Revenue, Quantity) |
| **Filters** | Page-level filters for the whole pivot (Status, Category) |

Values defaults to **Sum** — click the field → **Value Field Settings** to switch to Count, Average, Max.

## 3. Group dates, show shares

- Drag **OrderDate** to Rows → right-click a date → **Group** → tick **Quarters** and **Years**. (Works because you fixed the text-dates in Stage 1.)
- Right-click a value → **Show Values As → % of Grand Total** to see each region's share. Drag `Revenue` into Values **twice** for dollars and % side by side.

## 4. Slicers — clickable filters (and the dashboard trick)

**PivotTable Analyze → Insert Slicer** for `Region` and `Status` — big clickable buttons. The power move: one slicer can drive **several** pivots at once — right-click → **Report Connections** → tick the pivots. One click on "West" updates them all.

## 5. The rep leaderboard

The classic sales deliverable. Build a pivot of **SalesRep → Closed-Won revenue**, then:

1. Filter to Closed Won (Status in Filters, or a slicer).
2. Right-click a rep → **Filter → Top 10** → set to **Top 5**.
3. Right-click → **Sort → Largest to Smallest**.

That's your top-5 leaderboard. The solution file's `Summary` sheet mirrors it with `SUMIFS` and adds **data bars** for an instant visual ranking.

## 6. Refreshing

Pivots don't update on their own. When `Sales` changes, right-click → **Refresh** (or **Data → Refresh All**). Because `Sales` is a Table, new rows are picked up automatically on refresh.

## Practice

On the scratchpad: revenue by Region; Region × Quarter; Revenue by Category with a Status slicer; % of Grand Total; and the **Top 5 rep leaderboard**.

## Hand-off

You've got the numbers and the ranking. The last step is making them land for someone who has ten seconds and a meeting to run.

Take the [Stage 4 quiz](../quizzes/module-4), then finish with [Stage 5 — Present & capstone](05-present-capstone).
