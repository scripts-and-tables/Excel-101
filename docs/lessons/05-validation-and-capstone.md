---
layout: lesson
title: "Validation, what-if & capstone"
short_title: "Capstone"
subtitle: "Lock down inputs, run quick what-if analysis, and build the capstone dashboard."
module: 5
---

The final module covers the small touches that make a workbook trustworthy and reusable — and then you put everything together in a capstone dashboard.

## 1. Data validation — drop-downs

Data validation stops people typing `Wast` instead of `West`. On the `Quote Calculator` sheet of the working file:

1. Select the cell that should hold a region (e.g. `B3`).
2. **Data → Data Validation → Allow: List**.
3. Source: `Central,East,North,South,West` (or point at a range).
4. (Optional) Add an **Input Message** ("Pick a region") and an **Error Alert**.

For longer lists (e.g. all products), put the list on a hidden `Lists` sheet and reference the range:

```
Source: =Lists!$A$2:$A$30
```

## 2. Named ranges

Names make formulas readable. **Formulas → Define Name**.

- Name: `ProductList` → Refers to: `=Lists!$A$2:$A$30`
- Name: `SelectedProduct` → Refers to: `='Quote Calculator'!$B$4`

Now you can write `=XLOOKUP(SelectedProduct, Sales[Product], Sales[UnitPrice])` instead of opaque cell references.

The **Name Manager** (Ctrl + F3) lists everything you've named.

## 3. Goal Seek — what discount hits a target margin?

The `Quote Calculator` sheet computes `Margin %` based on Quantity, Discount, etc. Suppose you want to know: *what discount keeps the margin at exactly 30%?*

1. **Data → What-If Analysis → Goal Seek**.
2. **Set cell**: `B12` (Margin %)
3. **To value**: `0.30`
4. **By changing cell**: `B6` (Discount)

Excel solves it iteratively and writes the answer into `B6`. Useful for quoting, pricing, and sanity-checking deals.

## 4. Useful shortcuts for sales workflows

| Shortcut | What it does |
|----------|--------------|
| `Ctrl + T` | Create Excel Table |
| `Ctrl + Shift + L` | Toggle filters |
| `Ctrl + Shift + Down` | Select to bottom of column |
| `Ctrl + ;` | Insert today's date (static) |
| `Ctrl + Shift + 1` | Format as number with thousand separator |
| `Ctrl + Shift + 4` | Format as currency |
| `Ctrl + Shift + 5` | Format as percent |
| `Alt + =` | AutoSum the selection |
| `Alt + N + V` | Insert PivotTable |
| `F4` | Toggle absolute / relative reference (`A1` → `$A$1` → `A$1` → `$A1`) |

## 5. Capstone — build a sales dashboard

Now combine everything. Open `capstone.xlsx` (working). The brief:

> You are the new RevOps analyst. The VP of Sales wants a one-page dashboard for the management meeting on Monday. Build it on a sheet called **Dashboard**.

The dashboard must show:

1. **KPI strip** at the top:
   - Total revenue (Closed Won only)
   - Total order count (Closed Won)
   - Average order value
   - Refund rate (Refunded ÷ all orders)
2. **Revenue by Region** — bar or column chart.
3. **Monthly revenue trend** — line chart, full date range.
4. **Top 10 customers by revenue** — table with data bars.
5. **Revenue by Category × Region** — heat map (color scale).
6. **A slicer** for `Region` that filters at least one of the charts (use a pivot for that one).

Constraints:

- Only count `Status = "Closed Won"` for revenue figures.
- All currency formatted as `$1,234,567` (no decimals at this scale).
- Dashboard fits on one screen at 100% zoom.

When you're done, compare against `capstone.xlsx` (solution). Yours will look different — that's fine, there are many correct answers.

## You finished the course

Take the [Module 5 quiz](../quizzes/module-5) to wrap up.

If you want to go further from here, the natural next steps are:

- **Power Query** for connecting Excel directly to your CRM and refreshing data with one click.
- **Macros / VBA** for automating repetitive cleanup steps.
- **Power Pivot / DAX** for building data models across multiple tables.

Each of those is its own course — not in scope here.
