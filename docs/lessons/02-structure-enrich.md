---
layout: lesson
title: "Stage 2 — Structure & enrich"
short_title: "Structure & enrich"
subtitle: "Add the context the export is missing — pull in manager, quota and commission rate."
module: 2
---

> **Where we are:** Receive & clean → **Structure & enrich** → Explore & summarize → Pivot & rank → Present.

## Why enrich

Your cleaned `Sales` table knows *who* sold *what*, but not the facts you need to judge performance: each rep's **manager**, their **quota**, their **commission rate**. Those live in a separate reference table. **VLOOKUP** is how you bring them across — and it's the single most common formula you'll ever inherit, so you need to both write it and read it.

Open `module-2.xlsx` (working). Alongside `Sales` there's a small **`Reps`** table: `SalesRep · Region · Manager · AnnualQuota · CommissionRate`.

## 1. VLOOKUP in plain English

```
=VLOOKUP( what you're looking for,
          the table to look in,
          which column number to return,
          FALSE )
```

To stamp each order with its rep's manager:

```
=VLOOKUP([@SalesRep], Reps, 3, FALSE)
```

- **`[@SalesRep]`** — the value to find (this row's rep).
- **`Reps`** — the table to search; VLOOKUP matches against its **first column**.
- **`3`** — return column 3 (Manager). Use `4` for Quota, `5` for Commission Rate.
- **`FALSE`** — require an **exact** match. Always.

Add three columns — **Manager** (`3`), **Quota** (`4`), **Comm Rate** (`5`) — and your table is enriched. You'll use Quota and Comm Rate to compute attainment and commission in Stage 3.

## 2. Always pass FALSE (the big one)

The 4th argument is "approximate match?". Omit it or pass `TRUE` and VLOOKUP assumes the table is sorted and grabs the *nearest* value — silently returning **wrong answers**. For looking up a rep, a product, an ID — anything exact — you always want `FALSE`. Make it a reflex: **every VLOOKUP ends in `, FALSE)`**.

## 3. The six ways VLOOKUP breaks — and the fix

When VLOOKUP misbehaves it's almost always one of these (Microsoft's own list — see the solution's `Failure Modes` sheet):

| # | Symptom | Fix |
|---|---------|-----|
| 1 | Wrong row returned | You forgot `FALSE`. Always exact-match. |
| 2 | Can't look **left** of the key | VLOOKUP only looks right. Re-order columns, or use `XLOOKUP`. |
| 3 | Breaks when a column is inserted | The hard-coded `3` shifts. Reference a **Table** (`Reps`), or use `XLOOKUP`. |
| 4 | `#N/A` on values that look identical | One side is a **number stored as text** (Stage 1!). Fix with `VALUE`. |
| 5 | `#N/A` from invisible spaces | Trailing/leading spaces in the key. `TRIM` both sides. |
| 6 | `#N/A` because it isn't there | Wrap it: `=IFERROR(VLOOKUP(…), "unknown")`. |

That last one is the everyday one:

```
=IFERROR(VLOOKUP([@SalesRep], Reps, 3, FALSE), "unknown")
```

## 4. Modern note — XLOOKUP (Excel 365 / 2021+)

On a current Excel, **XLOOKUP** is the newer lookup Microsoft now recommends. It fixes most of the failures above by default — exact match, looks any direction, survives column inserts, built-in if-not-found:

```
=XLOOKUP([@SalesRep], Reps[SalesRep], Reps[Manager], "unknown")
```

We still teach VLOOKUP as the main event because it's in every legacy file you'll inherit, and XLOOKUP doesn't exist before Excel 2021. Know VLOOKUP cold; reach for XLOOKUP when everyone's on a modern version.

## Practice

In the `Exercises` sheet: add Manager / Quota / Comm Rate via VLOOKUP; try `TRUE` and watch it break; attempt a left-looking lookup; wrap one in `IFERROR`; (M365) redo with XLOOKUP.

## Hand-off

The file is now clean **and** enriched — every order carries its rep's manager, quota and commission rate. Time to actually answer some questions.

Take the [Stage 2 quiz](../quizzes/module-2), then continue to [Stage 3 — Explore & summarize](03-explore-summarize).
