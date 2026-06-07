---
layout: lesson
title: "Stage 3 — Explore & summarize"
short_title: "Explore & summarize"
subtitle: "Quick looks with filters, then the formulas that answer the everyday questions."
module: 3
---

> **Where we are:** Receive & clean → Structure & enrich → **Explore & summarize** → Pivot & rank → Present.

The highest-leverage stage. First you *explore* with filters and subtotals (no formulas), then you *summarize* with the conditional trio — `SUMIFS`, `COUNTIFS`, `AVERAGEIFS` — and build the sales KPIs you actually get asked for. Open `module-3.xlsx`; the `Sales` table is already enriched (it has `SalesManager` and `BrandManager`). Put each answer in column C of the `Exercises` sheet.

## Part A — Explore fast (no formulas)

### AutoFilter
Turn filters on with **Ctrl + Shift + L**. Filter `InvoiceType = "Sales"`, add `Area = "Deira"`, or `Category = "HPC"`. Filters **hide** rows; copying a filtered table copies only the visible rows.

### Read the Status Bar
Select the visible `SalesValue` cells and look bottom-right: Excel shows **Sum, Count, Average** instantly. Filter, select, read.

### SUBTOTAL — totals that respect the filter
`=SUM(…)` always adds every row; `SUBTOTAL` adds only the **visible** ones:

```
=SUBTOTAL(9,  M2:M2001)    // SUM of visible SalesValue
=SUBTOTAL(103, A2:A2001)   // COUNT of visible lines
=SUBTOTAL(101, M2:M2001)   // AVERAGE of visible
```

For automatic per-group totals, **Data → Subtotal** (sort by the group column first; works on a plain range, not a Table).

## Part B — Summarize with the conditional trio

The pattern is always: **the column to total, then pairs of (criteria column, criteria value)** — all conditions must be true.

```
=SUMIFS(Sales[SalesValue], Sales[Area], "Deira")
=SUMIFS(Sales[SalesValue], Sales[Area], "Deira", Sales[Category], "Food")
=COUNTIFS(Sales[InvoiceType], "Return")
=AVERAGEIFS(Sales[SalesValue], Sales[Category], "HPC")
```

> **Tip:** put the criteria in cells (`F1 = "Deira"`) and reference them — `=SUMIFS(Sales[SalesValue], Sales[Area], F1)` — so you change the filter by typing, not by rewriting formulas.

## Part C — Build the sales KPIs

Because **returns carry negative values**, plain `SUM` already gives you *net*:

```
Net sales     =SUM(Sales[SalesValue])
Gross sales   =SUMIFS(Sales[SalesValue], Sales[InvoiceType], "Sales")
Returns       =SUMIFS(Sales[SalesValue], Sales[InvoiceType], "Return")   (a negative number)
Return rate   =-SUMIFS(Sales[SalesValue], Sales[InvoiceType], "Return")
               / SUMIFS(Sales[SalesValue], Sales[InvoiceType], "Sales")
Avg sale      =AVERAGEIFS(Sales[SalesValue], Sales[InvoiceType], "Sales")
```

### Sales by manager — using the enriched columns
Stage 2 added `SalesManager` and `BrandManager` to every row, so now you can total by them directly:

```
=SUMIFS(Sales[SalesValue], Sales[BrandManager], "Imran Sheikh")   // a brand manager's net sales
=SUMIFS(Sales[SalesValue], Sales[SalesManager], "Tariq Aziz")     // a sales manager's team
```

### Quota attainment — clean × enrich × summarize together
A rep's net sales ÷ their quota, with the quota pulled from `Reps`:

```
=SUMIFS(Sales[SalesValue], Sales[SalesRep], "Mohammed Saleh")
 / VLOOKUP("Mohammed Saleh", Reps, 3, FALSE)
```

## Two mechanics you'll lean on

- **Fill down fast:** **double-click the fill handle** to copy a formula down the column. In a Table, typing one formula auto-fills it for you.
- **Absolute vs relative refs:** lock a cell that must not move with `$` (e.g. `$B$1`). Press **F4** while editing to cycle `A1 → $A$1 → A$1 → $A1`.

## Hand-off

You can answer almost any single question now. When you need to slice *many* ways at once — Area × Category × quarter — formulas get tedious. That's what PivotTables are for.

## Further learning — official Microsoft documentation

**Functions**
- [SUM](https://support.microsoft.com/office/043e1c7d-7726-4e80-8f32-07b23e057f89) · [AVERAGE](https://support.microsoft.com/office/047bac88-d466-426c-a32b-8f33eb960cf6) · [COUNT](https://support.microsoft.com/office/a59cd7fc-b623-4d93-87a4-d23bf411294c) · [COUNTA](https://support.microsoft.com/office/7dc98875-d5c1-46f1-9a82-53f3219e2509) · [MAX](https://support.microsoft.com/office/e0012414-9ac8-4b34-9a47-73e662c08098)
- [IF](https://support.microsoft.com/office/69aed7c9-4e8a-4755-a9bc-aa8bbff73be2) · [SUMIFS](https://support.microsoft.com/office/c9e748f5-7ea7-455d-9406-611cebce642b) · [COUNTIFS](https://support.microsoft.com/office/dda3dc6e-f74e-4aee-88bc-aa8c2a866842) · [AVERAGEIFS](https://support.microsoft.com/office/48910c45-1fc0-4389-a028-f7c5c3001690) · [SUBTOTAL](https://support.microsoft.com/office/7b027003-f060-4ade-9040-e478765b9939)

**Features & how-tos**
- [Filter data in a range or table](https://support.microsoft.com/en-us/office/filter-data-in-a-range-or-table-in-excel-01832226-31b5-4568-8806-38c37dcc180e) · [Insert subtotals in a list of data (Data > Subtotal)](https://support.microsoft.com/en-us/office/insert-subtotals-in-a-list-of-data-in-a-worksheet-7881d256-b4fa-4f81-b71e-b0a3d4a52b3a)
- [Use AutoSum to sum numbers](https://support.microsoft.com/en-us/office/use-autosum-to-sum-numbers-in-excel-543941e7-e783-44ef-8317-7d1bb85fe706) · [Switch between relative, absolute & mixed references](https://support.microsoft.com/en-us/office/switch-between-relative-absolute-and-mixed-references-dfec08cd-ae65-4f56-839e-5f0d8d0baca9)

Take the [Stage 3 quiz](../quizzes/module-3), then continue to [Stage 4 — Pivot & rank](04-pivot-rank).
