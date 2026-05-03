---
layout: lesson
title: "Pivot tables & pivot charts"
short_title: "Pivot tables"
subtitle: "The fastest way to slice a sales dataset — by region, period, rep, or product."
module: 3
---

Pivot tables are the fastest way to slice a sales dataset by region, period, rep, or product. Once you can build one, you can answer most "what's our revenue by X" questions in under a minute.

## 1. Build your first pivot

Open `module-3.xlsx`. Click any cell inside the `Sales` table, then **Insert → PivotTable → New Worksheet**.

In the PivotTable Fields pane:

- Drag **Region** to **Rows**
- Drag **Revenue** to **Values**

That's it — you have revenue by region.

## 2. The four drop zones

Every pivot table has four areas:

| Area | Use it for |
|------|------------|
| **Rows** | Categories you want listed down the side (Region, SalesRep, Product, Customer) |
| **Columns** | A second dimension across the top (often a time bucket — Quarter, Month) |
| **Values** | The numbers you want to summarise (Revenue, Quantity, Cost) |
| **Filters** | Page-level filters that apply to the whole pivot (Status, Category) |

Drag fields between zones to reshape the report instantly.

## 3. Group dates into months and quarters

Right-click any date in the Rows area → **Group**. Tick **Months** and **Years** (and optionally **Quarters**).

Your pivot now shows revenue by year → quarter → month, collapsible by clicking the +/- icons.

## 4. Show Values As — % of total, running total

By default, Values shows the sum. Right-click a value cell → **Show Values As** to change it:

- **% of Grand Total** — what share of total revenue does each region drive?
- **% of Parent Row Total** — within each year, what share is each quarter?
- **Running Total in** — cumulative revenue through the year.
- **Difference From** — month-over-month change.

You can drag the same field into Values **twice** — once as a sum, once as a percentage — to show both side by side.

## 5. Calculated fields — margin, commission, etc.

**PivotTable Analyze → Fields, Items & Sets → Calculated Field**.

- Name: `Margin`
- Formula: `=Revenue - Cost * Quantity`

Or commission:

- Name: `Commission`
- Formula: `=Revenue * 0.05`

The new field shows up like any other and respects all your slicers and filters.

> Heads up: calculated fields are computed on the *summed* values, not row-by-row, so for ratios like margin % you sometimes get more accurate results by adding a helper column to the source table instead.

## 6. Slicers and timelines

**PivotTable Analyze → Insert Slicer** gives you a clickable filter button. Add slicers for `Region`, `SalesRep`, and `Status` to make the pivot interactive.

**Insert → Timeline** does the same for date fields with a horizontal date scrubber — perfect for "show me Q3 only".

A single slicer can filter **multiple pivots at once** — right-click the slicer → **Report Connections** → tick the pivots you want it to control. This is the trick to building a real dashboard.

## 7. Pivot charts

With a pivot table selected, **PivotTable Analyze → PivotChart**. Pick a column or line chart. The chart updates live as you change the pivot — and slicers control the chart too.

For revenue trend, a **line chart** with Month on Rows and Revenue in Values is a sales classic.

## 8. Refreshing data

When the underlying `Sales` table changes (new orders added), pivots do **not** update automatically. Right-click any pivot → **Refresh**. Or, with multiple pivots, **Data → Refresh All** updates every pivot in the workbook.

If your `Sales` is a proper Excel Table (Module 1!), new rows are picked up automatically on refresh — no need to reset the source range.

## Practice

In the working file, build pivots for each of these on your own:

1. Total revenue by Region
2. Revenue by Region (rows) × Quarter (columns)
3. Top 5 SalesReps by revenue (use a Value Filter → Top 10 → set to Top 5)
4. Revenue by Category, with `Status` as a slicer
5. Monthly revenue trend pivot chart for one region (use a slicer)

Compare your results to the `Summary` sheet in the solution file (which mirrors what your pivots should show, built with `SUMIFS`).

Then take the [Module 3 quiz](../quizzes/module-3) and continue to [Module 4 — Visualizing sales data](04-visualization).
