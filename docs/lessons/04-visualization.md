---
layout: lesson
title: "Visualizing sales data"
short_title: "Visualization"
subtitle: "Pick the right chart, master conditional formatting, and use sparklines for trend signals."
module: 4
---

A chart's job is to make a number understandable in less than a second. Pick the wrong chart and you've made the data harder to read, not easier.

## 1. Pick the right chart for the question

| Question | Best chart |
|----------|-----------|
| Compare revenue across regions / reps / products | **Column** (vertical bars) or **Bar** (horizontal — better when labels are long) |
| Show how revenue changes over time | **Line** chart |
| Show two related metrics on the same chart (e.g. revenue + units) | **Combo** chart, with units on a secondary axis |
| Show share of total | **100% stacked bar** — avoid pie charts beyond 3–4 slices |
| Show distribution of order sizes | **Histogram** |
| Highlight a single KPI | **Big number** (just a styled cell) — no chart at all |

A good rule: if the user has to read a legend to understand what they're seeing, the chart isn't doing its job.

## 2. Build a column chart for revenue by region

In the working file, on the `Charts` sheet:

1. Build a small summary table (Region | Revenue) using `SUMIFS`.
2. Select it.
3. **Insert → Column or Bar Chart → Clustered Column**.

Clean it up:

- Click the chart title and rename it to something specific ("Revenue by Region — 2024–2025").
- Delete the gridlines and the legend (you only have one series).
- Format the value axis: right-click → **Format Axis** → **Number → Currency**, no decimals.
- Sort the source table descending on revenue so the bars are ranked.

## 3. Build a line chart for monthly trend

Build a `Month | Revenue` summary table, then **Insert → Line Chart**. Crucial for trend charts:

- Make sure the date axis is set to **Date axis** (not Text axis) so months are evenly spaced.
- If you have multiple years, consider showing them as separate lines (one per year, x-axis = month) for year-on-year comparison.

## 4. Combo charts: revenue and units

When you have two series on different scales (e.g. revenue in dollars, units sold), a single axis flattens one of them. Use a combo chart:

1. Select your data.
2. **Insert → Insert Combo Chart → Custom Combination**.
3. Set Revenue to **Clustered Column** with primary axis.
4. Set Units to **Line** with **secondary axis**.

## 5. Conditional formatting — the lightweight chart

Sometimes you don't need a chart at all. Conditional formatting on the table itself does the job:

- **Color scales** turn a region × month grid into a heat map. Select the values, **Home → Conditional Formatting → Color Scales**.
- **Data bars** put a mini bar inside the cell — perfect for ranking reps. **Home → Conditional Formatting → Data Bars**.
- **Top/Bottom rules** highlight the top 10 customers, or anything below average.
- **Custom rules with formulas** — e.g. highlight the entire row where `Status = "Refunded"`:
  - New Rule → Use a formula → `=$L2="Refunded"` → format with red fill.

## 6. Sparklines — tiny in-cell charts

A sparkline is a line/column chart that fits inside a single cell. Great for "rep performance over the last 12 months" columns.

1. **Insert → Sparkline → Line**.
2. Pick the data range (one row of monthly values per rep).
3. Pick where the sparkline should go.

Format options let you mark high points and low points.

## 7. Quick dashboard layout tips

- **Top-left = most important.** Put your headline KPI (e.g. total revenue) in the top-left corner. Eyes start there.
- **Group related charts.** Region performance charts together, time-trend charts together.
- **Consistent colours.** Pick one colour per region and use it everywhere — never let "West" be blue in one chart and orange in another.
- **Format numbers.** No raw `1234567.89`. Use `$1.2M` or `$1,234,567`. Right-click → Format Cells → Custom: `[>=1000000]"$"#,##0,,"M";[>=1000]"$"#,##0,"K";"$"#,##0`.

## Practice

Open the working file and build:

1. A **column chart** of revenue by region.
2. A **line chart** of monthly revenue.
3. A **combo chart** of revenue + units sold per month.
4. A **heat map** of revenue by Region × Quarter using a color scale.
5. A **sparkline** column showing each rep's monthly trend.

The solution file (`Monthly`, `Heatmap`, `RepBars` sheets) shows what each should roughly look like.

Take the [Module 4 quiz](../quizzes/module-4), then finish with [Module 5 — Validation, what-if & capstone](05-validation-and-capstone).
