---
layout: lesson
title: "Sales-power formulas"
short_title: "Formulas"
subtitle: "Logical, lookup, conditional aggregation, date, text, and dynamic-array formulas."
module: 2
---

This is the longest module — and the highest-leverage one. Master these formulas and 80% of your day-to-day sales reporting becomes a five-minute job.

The working file already has a `Sales` table with a `Revenue` column (`UnitPrice * Quantity * (1 - Discount)`). All examples below use that table.

## Logical: IF, IFS, AND/OR, IFERROR

### IF and IFS

`IF` is the workhorse:

```
=IF(Sales[@Revenue] > 1000, "Mid", "Small")
```

When you want more than two outcomes, use `IFS` (Excel 2019+) instead of nesting:

```
=IFS(
  Sales[@Revenue] > 5000, "Big",
  Sales[@Revenue] > 1000, "Mid",
  TRUE, "Small"
)
```

The final `TRUE` is the catch-all "else".

### AND / OR

Combine conditions:

```
=IF(AND(Sales[@Region]="West", Sales[@Status]="Closed Won"), "Bingo", "")
=IF(OR(Sales[@Status]="Refunded", Sales[@Status]="Cancelled"), "Bad", "OK")
```

### IFERROR / IFNA

Wrap formulas that might error so you get a clean output instead of `#N/A` or `#DIV/0!`:

```
=IFERROR(A2/B2, 0)
=IFNA(XLOOKUP(...), "not found")
```

> Note: with `XLOOKUP` you can use the built-in `if_not_found` argument and skip `IFNA` entirely.

## Lookup: XLOOKUP (and a brief INDEX/MATCH note)

### XLOOKUP — the modern default

```
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])
```

Examples:

```
=XLOOKUP("Headset Pro", Sales[Product], Sales[UnitPrice], "not found")
=XLOOKUP(B2, Sales[OrderID], Sales[Customer], "missing")
```

Why XLOOKUP beats VLOOKUP:

| | VLOOKUP | XLOOKUP |
|---|---------|---------|
| Look left | ❌ | ✅ |
| Default match | Approximate (dangerous) | Exact (safe) |
| Built-in error handler | ❌ — wrap in IFERROR | ✅ `if_not_found` |
| Survives column inserts | ❌ — column index breaks | ✅ — references the column directly |

### INDEX / MATCH (when you need it)

INDEX/MATCH still has its place — older versions of Excel, complex 2D lookups, performance on huge files:

```
=INDEX(Sales[UnitPrice], MATCH("Headset Pro", Sales[Product], 0))
```

### VLOOKUP (legacy mention)

You will still inherit files that use `VLOOKUP`. For reference:

```
=VLOOKUP(lookup_value, table, column_number, FALSE)
```

The `FALSE` is critical — without it, VLOOKUP does an approximate match and silently returns wrong answers. **For new work, prefer XLOOKUP.**

## Conditional aggregation: SUMIFS, COUNTIFS, AVERAGEIFS, MAXIFS

These are the formulas you will use *every day* when reporting sales.

```
=SUMIFS(Sales[Revenue], Sales[Region], "West")
=SUMIFS(Sales[Revenue], Sales[Region], "West", Sales[Status], "Closed Won")
=COUNTIFS(Sales[Status], "Closed Won")
=AVERAGEIFS(Sales[Revenue], Sales[Category], "Hardware")
=MAXIFS(Sales[Revenue], Sales[Category], "Software")
```

The pattern is always: **sum/count column, then condition column + condition value**, repeated.

> Tip: put your criteria values in cells (e.g. `B1 = "West"`) and reference them — `SUMIFS(Sales[Revenue], Sales[Region], B1)` — so you can change the filter without rewriting formulas.

## Date functions for sales

```
=TODAY()                                    // Today's date, refreshes daily
=EOMONTH(TODAY(), 0)                        // End of current month
=EOMONTH(TODAY(), -1) + 1                   // First of current month
=DATEDIF(Sales[@OrderDate], TODAY(), "d")   // Age of order in days
=YEAR(Sales[@OrderDate])                    // Pull out the year
=NETWORKDAYS(Sales[@OrderDate], TODAY())    // Working days since order
```

Common combo — month label for a pivot:

```
=TEXT(Sales[@OrderDate], "yyyy-mm")
```

## Text functions for cleaning IDs and names

```
=LEFT("ACME-2024-0042", 4)        // "ACME"
=RIGHT("ACME-2024-0042", 4)       // "0042"
=MID("ACME-2024-0042", 6, 4)      // "2024"
=TRIM("  Acme  Corp  ")           // "Acme Corp" (also collapses internal doubles)
=SUBSTITUTE("Acme,Corp", ",", " ") // "Acme Corp"
=CONCAT("Q", QUARTER(...))         // (use TEXTJOIN instead, see below)
=TEXTJOIN(" - ", TRUE, "ACME", 2024, 42)  // "ACME - 2024 - 42"
=TEXT(1234.5, "$#,##0.00")        // "$1,234.50"
```

## Dynamic arrays: UNIQUE, SORT, FILTER

These are recent additions (Excel 2021 / Microsoft 365) and they are *transformative* for sales reporting.

```
=UNIQUE(Sales[Region])                              // List each region once
=SORT(UNIQUE(Sales[SalesRep]))                       // Same, sorted
=FILTER(Sales, Sales[SalesRep]="Anna Becker")        // All of Anna's orders
=FILTER(Sales, (Sales[Region]="West") * (Sales[Status]="Closed Won"))  // Multi-condition
=SORT(FILTER(Sales, Sales[Status]="Closed Won"), 13, -1)  // Sorted by Revenue desc
```

The result "spills" into adjacent cells — you write one formula, Excel fills the range.

## Practice

The working file has an `Exercises` sheet with 12 prompts. Try each one and compare to the solution file. Then take the [Module 2 quiz](../quizzes/module-2).

Move on to [Module 3 — Pivot tables](03-pivot-tables).
