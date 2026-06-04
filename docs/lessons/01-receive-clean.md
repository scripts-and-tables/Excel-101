---
layout: lesson
title: "Stage 1 — Receive & clean"
short_title: "Receive & clean"
subtitle: "A CRM export just landed in your inbox. Make it something you can trust."
module: 1
---

> **The workflow:** Receive & clean → Structure & enrich → Explore & summarize → Pivot & rank → Present. We'll carry one sales export through all five. This is Stage 1.

## Why we start here

A colleague drops a CSV from the CRM on your desk: "Can you tell me how we did?" It's *usable*, but rough — extra spaces in customer names, a few regions in CAPS, duplicate rows, numbers that arrived as text and won't add up. **Every number you produce later is only as good as this step.** So before any analysis, you make the file trustworthy.

Open `module-1.xlsx` (working). The `RawData` sheet is a messy export, on purpose.

## 1. Turn the export into an Excel Table

Click any cell, press **Ctrl + T**, confirm "My table has headers" → **OK**. A Table gives you auto-expanding ranges, **structured references** (`Sales[Revenue]` instead of `M2:M2003`), filter buttons, and banded rows. Rename it under **Table Design → Table Name →** `Sales`. From here on, the data *is* `Sales`, and it travels with us to the end of the course.

## 2. Trim spaces — `TRIM`

`" Acme Corp "` and `"Acme Corp"` look identical but count as two customers. In a helper column:

```
=TRIM(Sales[@Customer])
```

`TRIM` also collapses double spaces inside the text.

## 3. Strip hidden characters — `CLEAN`

Exports carry invisible junk — line breaks, tabs — pasted from other systems. They break lookups and grouping later. Strip them:

```
=CLEAN(Sales[@Customer])
=TRIM(CLEAN(Sales[@Customer]))     // the belt-and-braces combo
```

## 4. Fix repeated typos — Find & Replace

Same mistake on many rows (`NORTH` instead of `North`, an old product name)? Don't edit cell by cell — **Ctrl + H**: *Find* `NORTH` → *Replace with* `North` → **Replace All**. Tick **Match entire cell contents** for exact-only swaps. (**Ctrl + F** is find-only.)

## 5. Fix numbers stored as text — `VALUE`

A column that *looks* numeric but won't `SUM` arrived as **text**. Tells: values sit on the **left** of the cell, with a green warning triangle. Convert:

```
=VALUE(Sales[@Quantity])
```

…or select the column → click ⚠️ → **Convert to Number**. Until you do, `SUM` silently under-counts.

## 6. Fix dates stored as text — `DATEVALUE`

`"2024-03-15"` as text can't be sorted, grouped by month, or used in date math (you'll need all three in Stage 4). Convert with `=DATEVALUE(Sales[@OrderDate])`, then format as a date (**Ctrl + Shift + #**). Real dates right-align; text dates left-align.

## 7. Split and combine text — `LEFT` / `MID` / `RIGHT` / `SUBSTITUTE` / `&`

Real exports bury data inside codes. Pull it apart and put it back together:

```
=LEFT("ORD-2024-0042", 3)              // "ORD"
=MID("ORD-2024-0042", 5, 4)            // "2024"
=RIGHT("ORD-2024-0042", 4)             // "0042"
=SUBSTITUTE("Acme;Corp;LLC", ";", " ") // swap a separator
="ORD-" & A2                           // join text with &
```

The `Cleanup Practice` sheet has a cell for each of these.

## 8. Lock it in — Paste Special

A cleanup formula still points at the messy original. Before deleting that column, freeze the results: select the helper → **Ctrl + C** → **Paste Special** (**Ctrl + Alt + V**) → **Values**. Paste Special also does **Transpose** (flip rows/columns) and **Operations** (e.g. multiply a column by an FX rate in place).

## 9. Dedupe, sort, freeze

- **Data → Remove Duplicates** — untick `OrderID`, tick the columns that together identify an order. (Two deliberate dupes here.)
- **Data → Sort** — Region A→Z, then OrderDate newest-first.
- **View → Freeze Top Row** — keep headers visible.

## 10. Format numbers for reading (display ≠ value)

Formatting changes how a number *looks*, not what it *is*. Quick keys after selecting cells: **Ctrl+Shift+4** currency · **Ctrl+Shift+5** percent · **Ctrl+Shift+1** thousands · **Ctrl+Shift+3** date. For the compact dashboard look, **Ctrl + 1 → Custom**:

```
[>=1000000]"$"#,##0,,"M";[>=1000]"$"#,##0,"K";"$"#,##0     →  $1.2M / $340K / $820
```

## Hand-off

Your `Sales` Table is now clean, typed, deduped, sorted, and readable. But the export is missing context — who each rep's *manager* is, their *quota*, their *commission rate*. That's the next step.

Take the [Stage 1 quiz](../quizzes/module-1), then continue to [Stage 2 — Structure & enrich](02-structure-enrich).
