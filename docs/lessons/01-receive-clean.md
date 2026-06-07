---
layout: lesson
title: "Stage 1 — Receive & clean"
short_title: "Receive & clean"
subtitle: "A supermarket sales export just landed. Make it something you can trust."
module: 1
---

> **The workflow:** Receive & clean → Structure & enrich → Explore & summarize → Pivot & rank → Present. We'll carry one Dubai FMCG sales export through all five. This is Stage 1.

## Why we start here

You pull a sales export from the system: every line is one invoice — a brand sold to a Dubai supermarket branch, with a quantity and an AED value. It's *usable*, but rough: extra spaces in customer names, an area typed in CAPS, a duplicate line, a value that came through as text and won't add up. **Every number you produce later is only as good as this step.**

Open `module-1.xlsx` (working). The `RawData` sheet is a messy export, on purpose. Its columns:

`OrderNumber · Date · InvoiceType · CustomerCode · Customer · BranchCode · Branch · Area · SalesRep · Brand · Category · SalesQuantity · SalesValue`

A few things to know about this data:
- **InvoiceType** is `Sales` or `Return`. **Returns carry negative** quantity *and* value — that's how a refund/return shows up.
- **Customer** is the supermarket chain (Carrefour, Lulu, Spinneys…); **Branch** is that chain in a Dubai **Area** (e.g. `Lulu Hypermarket - Al Barsha`).

## 1. Turn the export into an Excel Table

Click any cell, press **Ctrl + T**, confirm "My table has headers" → **OK**. A Table gives you auto-expanding ranges, **structured references** (`Sales[SalesValue]` instead of `M2:M2003`), filter buttons and banded rows. Rename it under **Table Design → Table Name →** `Sales`. It travels with us to the end of the course.

## 2. Trim spaces — `TRIM`

`" Carrefour "` and `"Carrefour"` count as two different customers in a pivot. In a helper column:

```
=TRIM(Sales[@Customer])
```

`TRIM` also collapses double spaces inside the text.

## 3. Strip hidden characters — `CLEAN`

Exports carry invisible junk — line breaks, tabs — from other systems. Strip them:

```
=CLEAN(Sales[@Customer])
=TRIM(CLEAN(Sales[@Customer]))     // belt-and-braces
```

## 4. Fix repeated typos — Find & Replace

When the same mistake repeats — an area as `DEIRA`, a misspelled chain — fix it in bulk with **Ctrl + H**: *Find* `DEIRA` → *Replace with* `Deira` → **Replace All**. Tick **Match entire cell contents** for exact-only swaps. (**Ctrl + F** is find-only.)

## 5. Fix numbers stored as text — `VALUE`

A `SalesValue` (or `SalesQuantity`) column that *looks* numeric but won't `SUM` arrived as **text**. Tells: values sit on the **left** of the cell, with a green warning triangle. Convert:

```
=VALUE(Sales[@SalesValue])
```

…or select the column → click ⚠️ → **Convert to Number**. Until you do, `SUM` silently under-counts.

## 6. Fix dates stored as text — `DATEVALUE`

A `Date` stored as text can't be sorted, grouped by month, or used in date math (you'll group by quarter in Stage 4). Convert with `=DATEVALUE(Sales[@Date])`, then format as a date (**Ctrl + Shift + #**). Real dates right-align; text dates left-align.

## 7. Split and combine text — `LEFT` / `MID` / `RIGHT` / `SUBSTITUTE` / `&`

Codes often pack several facts together. Pull them apart and rejoin:

```
=LEFT("SO-100562", 2)               // "SO"
=MID("SO-100562", 4, 6)             // "100562"
=RIGHT("SO-100562", 6)              // "100562"
=SUBSTITUTE("Deira;Dubai", ";", " ") // swap a separator
="SO-" & A2                          // join text with &
```

The `Cleanup Practice` sheet has a cell for each.

## 8. Lock it in — Paste Special

A cleanup formula still points at the messy original. Before deleting that column, freeze the results: select the helper → **Ctrl + C** → **Paste Special** (**Ctrl + Alt + V**) → **Values**. Paste Special also does **Transpose** and **Operations** (e.g. multiply a column by a cell in place).

## 9. Dedupe, sort, freeze

- **Data → Remove Duplicates** — untick `OrderNumber`, tick the columns that together identify a line. (Two deliberate dupes here.)
- **Data → Sort** — Customer A→Z, then Date newest-first.
- **View → Freeze Top Row.**

## 10. Format numbers for reading (display ≠ value)

Formatting changes how a number *looks*, not what it *is*. Quick keys: **Ctrl+Shift+1** thousands · **Ctrl+Shift+3** date. For AED money, **Ctrl + 1 → Custom**:

```
"AED" #,##0          →  AED 1,250
"AED" #,##0.00       →  AED 1,250.50
```

## Hand-off

Your `Sales` Table is now clean, typed, deduped, sorted and readable. But the export is missing context — each rep's **manager** and **quota**, and each brand's **brand manager**. That's next.

## Further learning

**Official Microsoft docs**
- [Create and format an Excel table](https://support.microsoft.com/office/create-and-format-tables-e81aa349-b006-4f8a-9806-5af9df0ac664)
- [Top ten ways to clean your data](https://support.microsoft.com/en-us/office/top-ten-ways-to-clean-your-data-2844b620-677c-47a7-ac3e-c2e157d1db19)
- [TRIM](https://support.microsoft.com/office/410388fa-c5df-49c6-b16c-9e5630b479f9) · [CLEAN](https://support.microsoft.com/office/26f3d7c5-475f-4a9c-90e5-4b8ba987ba41) · [VALUE](https://support.microsoft.com/office/257d0108-07dc-437d-ae1c-bc2d3953d8c2) · [DATEVALUE](https://support.microsoft.com/office/df8b07d4-7761-4a93-bc33-b7471bbff252)
- [LEFT](https://support.microsoft.com/office/9203d2d2-7960-479b-84c6-1ea52b99640c) · [MID](https://support.microsoft.com/office/d5f9e25c-d7d6-472e-b568-4ecb12433028) · [RIGHT](https://support.microsoft.com/office/240267ee-9afa-4639-a02b-f19e1786cf2f) · [SUBSTITUTE](https://support.microsoft.com/office/6434944e-a904-4336-a9b0-1e58df3bc332)
- [Find or replace text and numbers](https://support.microsoft.com/en-us/office/find-or-replace-text-and-numbers-on-a-worksheet-0e304ca5-ecef-4808-b90f-fdb42f892e90) · [Find and remove duplicates](https://support.microsoft.com/en-us/office/find-and-remove-duplicates-00e35bea-b46a-4d5d-b28e-66a552dc138d)
- [Move or copy cells (Paste Special)](https://support.microsoft.com/en-us/office/move-or-copy-cells-rows-and-columns-3ebbcafd-8566-42d8-8023-a2ec62746cfc) · [Available number formats](https://support.microsoft.com/en-us/office/available-number-formats-in-excel-0afe8f52-97db-41f1-b972-4b46e9f1e8d2)

**Video:** [Cleaning data & Excel Tables — video tutorials](https://www.youtube.com/results?search_query=excel+clean+data+tables+trim+tutorial)

Take the [Stage 1 quiz](../quizzes/module-1), then continue to [Stage 2 — Structure & enrich](02-structure-enrich).
