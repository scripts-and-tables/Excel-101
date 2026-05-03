---
layout: lesson
title: "Working with sales data"
short_title: "Tables & data"
subtitle: "Tables, structured references, sorting, filtering, and basic cleanup."
module: 1
---

## Why this matters for sales

Most sales workbooks start as a CSV export from your CRM or order system. The data is usable, but rough — extra spaces in customer names, mixed casing, the odd duplicate row. If you skip the cleanup step, every formula and pivot you build downstream is one bad row away from giving you the wrong number.

This module walks through the cleanup + structuring steps you should run on **every** new sales export before you analyse anything.

## 1. Convert raw data into an Excel Table

Open `module-1.xlsx`, select any cell inside the data, then press **Ctrl + T**. Confirm "My table has headers" and click OK.

Tables give you four big wins:

1. Auto-expanding ranges — adding a new row at the bottom is automatically included in formulas.
2. **Structured references** like `Sales[Revenue]` instead of `M2:M2003`.
3. Built-in filter buttons.
4. Banded rows for readability.

Rename the table: with any cell selected inside the table, go to **Table Design → Table Name** and call it `Sales`.

## 2. Trim spaces from text columns

A few of the customer names have leading/trailing spaces (compare `"Acme Corp"` vs `" Acme Corp "`). Pivots will treat those as two different customers.

In a helper column, use:

```
=TRIM(Sales[@Customer])
```

Then **Copy → Paste Special → Values** back over the original column and delete the helper.

## 3. Standardise text casing

A few `Region` values are uppercase (`NORTH`) instead of `North`. Use:

```
=PROPER(Sales[@Region])
```

Same trick — paste values back, delete helper. (For two-letter codes like `US` or `EU` use `UPPER` instead.)

## 4. Remove duplicates

Select the table, then **Data → Remove Duplicates**. Untick `OrderID` (since duplicates often have a fresh ID) and tick the columns that *together* uniquely identify an order — typically `OrderDate`, `Customer`, `Product`, `Quantity`.

Excel will tell you how many duplicates it removed. The working file has two deliberate duplicates.

## 5. Multi-level sort

**Data → Sort** lets you stack sort levels. A useful default for sales data:

1. `Region` — A to Z
2. `OrderDate` — Newest to Oldest

Now scrolling the table is a tour through each region's recent activity.

## 6. Freeze panes & view tricks

- **View → Freeze Top Row** keeps headers visible while scrolling.
- **View → Split** lets you scroll two parts of the same sheet side by side — useful when you're comparing the top of a long table to a row near the bottom.
- **Ctrl + Shift + L** toggles AutoFilter on/off.

## 7. Quick AutoFilter recipes

Click the filter arrow on any column header.

- **Region → tick only "West"** to focus on one region.
- **Status → tick everything except "Refunded"** to exclude refunds from your view.
- **OrderDate → Date Filters → This Month** for a quick MTD view.

Filters do **not** delete rows — they just hide them. Anything you copy from a filtered table copies only the visible rows.

## Recap

You now have:

- A properly-named **Excel Table** (`Sales`)
- Cleaned customer names and region casing
- No duplicates
- A sensible sort order

This is the foundation everything else in the course builds on. Take the [Module 1 quiz](../quizzes/module-1), then move on to [Module 2 — Sales-power formulas](02-formulas).
