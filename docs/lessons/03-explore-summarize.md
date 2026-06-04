---
layout: lesson
title: "Stage 3 — Explore & summarize"
short_title: "Explore & summarize"
subtitle: "Quick looks first, then the formulas that answer the everyday questions."
module: 3
---

> **Where we are:** Receive & clean → Structure & enrich → **Explore & summarize** → Pivot & rank → Present.

This is the highest-leverage stage. First you *explore* with filters and subtotals (no formulas), then you *summarize* with the conditional trio — `SUMIFS`, `COUNTIFS`, `AVERAGEIFS` — and build the sales KPIs your manager actually asks for. Open `module-3.xlsx`; put each answer in column C of the `Exercises` sheet.

## Part A — Explore fast (no formulas)

### AutoFilter
Turn filters on with **Ctrl + Shift + L**. Filter `Status = "Closed Won"`, add `Region = "West"`, or `Number Filters → Top 10`. Filters **hide** rows; copying a filtered table copies only the visible rows.

### Read the Status Bar
Select the visible `Revenue` cells and look bottom-right: Excel shows **Sum, Count, Average** of the selection instantly. The fastest "what's the number?" in Excel — filter, select, read.

### SUBTOTAL — totals that respect the filter
`=SUM(…)` always adds every row; `SUBTOTAL` adds only the **visible** ones:

```
=SUBTOTAL(9,  M2:M2001)    // SUM of visible rows   (109 also ignores hidden)
=SUBTOTAL(103, A2:A2001)   // COUNT of visible rows
=SUBTOTAL(101, M2:M2001)   // AVERAGE of visible rows
```

For automatic per-group totals, **Data → Subtotal** (sort by the group column first; it works on a plain range, not a Table).

## Part B — Summarize with the conditional trio

The formulas you'll use **every day**. The pattern is always: **the column to total, then pairs of (criteria column, criteria value)** — all conditions must be true.

```
=SUMIFS(Sales[Revenue], Sales[Region], "West")
=SUMIFS(Sales[Revenue], Sales[Region], "West", Sales[Status], "Closed Won")
=COUNTIFS(Sales[Status], "Closed Won")
=AVERAGEIFS(Sales[Revenue], Sales[Category], "Hardware")
```

> **Tip:** put the criteria in cells (`F1 = "West"`) and reference them — `=SUMIFS(Sales[Revenue], Sales[Region], F1)` — so you change the filter by typing, not by rewriting formulas.

`IF` tags rows by a rule, so you can filter/pivot on the tag later:

```
=IF([@Revenue] > 2000, "Large", "Standard")
```

## Part C — Build the sales KPIs

KPIs are just these formulas combined:

```
Win rate    =COUNTIFS(Sales[Status],"Closed Won") / COUNTA(Sales[OrderID])
AOV         =SUMIFS(Sales[Revenue],Sales[Status],"Closed Won")
             / COUNTIFS(Sales[Status],"Closed Won")
Refund rate =COUNTIFS(Sales[Status],"Refunded") / COUNTA(Sales[OrderID])
```

### Commission — the sales-specific one
Stage 2 gave every rep a commission rate. A rep's commission is their Closed-Won revenue × that rate, and the rate comes straight back out of `Reps` with VLOOKUP:

```
=SUMIFS(Sales[Revenue], Sales[SalesRep], "Anna Becker", Sales[Status], "Closed Won")
 * VLOOKUP("Anna Becker", Reps, 5, FALSE)
```

That single formula is clean → enrich → summarize, all three stages working together.

## Two mechanics you'll lean on

- **Fill down fast:** **double-click the fill handle** (bottom-right of the cell) to copy a formula down the whole column. In a Table, typing one formula auto-fills the column for you.
- **Absolute vs relative refs:** lock a cell that must not move with `$` — `=SUMIFS(Sales[Revenue], Sales[Region], $H1) * $B$1`. Press **F4** while editing to cycle `A1 → $A$1 → A$1 → $A1`.

## Hand-off

You can now answer almost any single question. But when you need to slice *many* ways at once — by region *and* quarter *and* rep — formulas get tedious. That's what PivotTables are for.

Take the [Stage 3 quiz](../quizzes/module-3), then continue to [Stage 4 — Pivot & rank](04-pivot-rank).
