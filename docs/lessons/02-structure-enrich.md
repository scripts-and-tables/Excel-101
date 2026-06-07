---
layout: lesson
title: "Stage 2 — Structure & enrich"
short_title: "Structure & enrich"
subtitle: "Add the context the export is missing — sales managers, quotas and brand managers."
module: 2
---

> **Where we are:** Receive & clean → **Structure & enrich** → Explore & summarize → Pivot & rank → Present.

## Why enrich

Your cleaned `Sales` table knows *who sold what to whom*, but not the facts you need to judge performance: each rep's **sales manager** and **annual quota**, and each brand's **brand manager**. Those live in two small reference tables. **VLOOKUP** brings them across — and it's the single most common formula you'll ever inherit.

Open `module-2.xlsx` (working). Alongside `Sales` you have:
- **`Reps`** — `SalesRep · Manager · AnnualQuota` *(a rep covers customers/areas)*
- **`Brands`** — `Brand · Category · BrandManager` *(a brand manager covers brands)*

## 1. VLOOKUP in plain English

```
=VLOOKUP( what you're looking for,
          the table to look in,
          which column number to return,
          FALSE )
```

Stamp each line with its rep's **sales manager**:

```
=VLOOKUP([@SalesRep], Reps, 2, FALSE)
```

- **`[@SalesRep]`** — the value to find (this row's rep).
- **`Reps`** — the table to search; VLOOKUP matches its **first column**.
- **`2`** — return column 2 (Manager). Use `3` for AnnualQuota.
- **`FALSE`** — exact match. Always.

Then bring in the **brand manager** from the other table:

```
=VLOOKUP([@Brand], Brands, 3, FALSE)
```

Now every line carries `SalesManager`, `Quota`, and `BrandManager`. In Stage 3 you'll use these to get *net sales by brand manager* and *quota attainment by rep*.

## 2. Always pass FALSE (the big one)

The 4th argument is "approximate match?". Omit it or pass `TRUE` and VLOOKUP assumes the table is sorted and grabs the *nearest* value — silently returning **wrong answers**. For a rep, a brand, a code — anything exact — always use `FALSE`. Make it a reflex: **every VLOOKUP ends in `, FALSE)`**.

## 3. The six ways VLOOKUP breaks — and the fix

Almost every VLOOKUP problem is one of these (Microsoft's own list — see the solution's `Failure Modes` sheet):

| # | Symptom | Fix |
|---|---------|-----|
| 1 | Wrong row returned | You forgot `FALSE`. Always exact-match. |
| 2 | Can't look **left** of the key | VLOOKUP only looks right. Re-order columns, or use `XLOOKUP`. |
| 3 | Breaks when a column is inserted | The hard-coded `2`/`3` shifts. Reference a **Table** (`Reps`/`Brands`), or use `XLOOKUP`. |
| 4 | `#N/A` on values that look identical | One side is a **number stored as text** (Stage 1!). Fix with `VALUE`. |
| 5 | `#N/A` from invisible spaces | Trailing/leading spaces in the key. `TRIM` both sides. |
| 6 | `#N/A` because it isn't there | Wrap it: `=IFERROR(VLOOKUP(…), "unknown")`. |

The everyday one:

```
=IFERROR(VLOOKUP([@Brand], Brands, 3, FALSE), "unknown")
```

## 4. Modern note — XLOOKUP (Excel 365 / 2021+)

On a current Excel, **XLOOKUP** is the newer lookup Microsoft recommends — exact match by default, looks any direction, survives column inserts, built-in if-not-found:

```
=XLOOKUP([@Brand], Brands[Brand], Brands[BrandManager], "unknown")
```

We still teach VLOOKUP as the main event: it's in every legacy file you'll inherit, and XLOOKUP doesn't exist before Excel 2021.

## Practice

In the `Exercises` sheet: add `SalesManager`, `Quota` (from `Reps`) and `BrandManager` (from `Brands`); try `TRUE` and watch it break; attempt a left-looking lookup; wrap one in `IFERROR`; (M365) redo with XLOOKUP.

## Hand-off

The file is clean **and** enriched — every line now carries its sales manager, quota and brand manager. Time to answer questions.

## Further learning — official Microsoft documentation

**Functions**
- [VLOOKUP](https://support.microsoft.com/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1) · [IFERROR](https://support.microsoft.com/en-us/office/iferror-function-c526fd07-caeb-47b8-8bb6-63f3e417f611) · [XLOOKUP (modern alternative)](https://support.microsoft.com/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929)
- [Look up values with VLOOKUP, INDEX or MATCH](https://support.microsoft.com/en-us/office/look-up-values-with-vlookup-index-or-match-68297403-7c3c-4150-9e3c-4d348188976b) · [MATCH](https://support.microsoft.com/office/e8dffd45-c762-47d6-bf89-533f4a37673a)

**Guides**
- [VLOOKUP troubleshooting tips (the six failure modes)](https://support.microsoft.com/en-us/office/quick-reference-card-vlookup-troubleshooting-tips-6fe7fe1b-709b-4958-adfb-9f2a409dcf38)

Take the [Stage 2 quiz](../quizzes/module-2), then continue to [Stage 3 — Explore & summarize](03-explore-summarize).
