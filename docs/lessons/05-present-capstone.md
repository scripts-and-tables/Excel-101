---
layout: lesson
title: "Stage 5 — Present & capstone"
short_title: "Present & capstone"
subtitle: "Make the numbers land on one page — then run the whole workflow yourself."
module: 5
---

> **Where we are:** Receive & clean → Structure & enrich → Explore & summarize → Pivot & rank → **Present**.

You've done the analysis. The last step is turning it into something the VP reads in ten seconds: a clean one-pager. Open `module-5.xlsx` and build it on the `One-Pager` sheet.

## 1. A KPI strip

Put the headline numbers across the top, big and formatted — they're the Stage 3 KPIs:

- Net sales · Gross sales · Return rate · Average sale value · # invoices

AED as `AED 1,250,000`, return rate as `%`. Top-left is where eyes land — put the most important number there.

## 2. One clean chart

You rarely need more than one. Build an `Area | Net Sales` summary, select it, **Insert → Column Chart**. Then: rename the title, delete gridlines and the legend, format the axis as AED, and **sort the source descending** so the bars rank themselves.

## 3. Let conditional formatting do the flagging

On a summary table or the data itself (**Home → Conditional Formatting**):

- **Color scales** turn a Category × Area grid into a heat map.
- **Data bars** rank brands or reps in-cell.
- **Icon sets** add traffic lights (e.g. on quota attainment).
- **Highlight → Duplicate Values** flags repeated customers/IDs — a fast check on Stage 1's dedupe.
- **Formula rule:** New Rule → *Use a formula* → `=$C2="Return"` → red fill, and every **Return** line's whole row turns red.

Manage it all from **Conditional Formatting → Manage Rules**.

## 4. Layout tips

- **Top-left = most important.**
- **One colour per area** everywhere; never show a raw `1234567.89` — use `AED 1.2M` or `AED 1,234,567`.
- Fit it on **one screen at 100%**.

## Keyboard shortcut cheat-sheet

| Shortcut | Does |
|----------|------|
| `Ctrl + T` | Create a Table |
| `Ctrl + Shift + L` | Toggle AutoFilter |
| `Ctrl + Alt + V` | Paste Special (Values, Transpose…) |
| `Ctrl + H` / `Ctrl + F` | Replace / Find |
| `Ctrl + D` | Fill down |
| `Ctrl + Shift + ↓` | Select to bottom of column |
| `Ctrl + 1` | Format Cells |
| `Ctrl + Shift + 1` | Number format (thousands) |
| `Alt + =` | AutoSum |
| `Alt + N + V` | Insert PivotTable |
| `F2` | Edit the active cell |
| `F4` | Toggle absolute/relative reference |

---

## 🎯 Capstone — run the whole workflow

Open `capstone.xlsx` (working). It's a **fresh sales file** plus the `Reps` and `Brands` tables — same shape you started with, new data. Your job is the whole workflow, end to end.

> **Brief:** You're the new analyst on the Dubai account. Before Monday's meeting, answer the questions on the `Questions` sheet (column C), then drop the headline numbers into a one-page summary. Use PivotTables, AutoFilter, `SUMIFS`/`COUNTIFS`, and `VLOOKUP` — whatever's fastest.

The 11 questions (on the `Questions` sheet):

1. **Net sales** (sum of SalesValue).
2. **Gross sales** (Sales invoices only).
3. **Total returns** value.
4. **Return rate** (returns ÷ gross).
5. **Number of invoice lines.**
6. **Average sale value** (Sales lines only).
7. Net sales in the **Deira** area.
8. Net sales for the **Food** category.
9. Net sales for customer **Carrefour**.
10. Net sales for brand manager **Imran Sheikh** (use the `BrandManager` column).
11. **Mohammed Saleh's quota attainment** (his net sales ÷ his quota from `Reps`).

**Ground rules:** returns are negative, so `SUM` gives net automatically; format AED as `AED 1,234,567` and rates as `%`; prefer a **live formula** over a typed number so it survives a data refresh.

Compare against `capstone.xlsx` (solution) — the `Answers` sheet has a working formula for each. There's often more than one correct route.

## You finished the course

You can take a raw Dubai sales export and walk it through the whole process: **clean → enrich → explore & summarize → pivot & rank → present**. That's the 80% of Excel a sales team uses, in the order you'd actually use it.

## Further learning

**Official Microsoft docs**
- [Add, change, or clear conditional formats](https://support.microsoft.com/office/add-change-or-clear-conditional-formats-fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f)
- [Create a chart from start to finish](https://support.microsoft.com/en-us/office/create-a-chart-from-start-to-finish-0baf399e-dd61-4e18-8a73-b3fd5d5680c2)
- [Available number formats in Excel](https://support.microsoft.com/en-us/office/available-number-formats-in-excel-0afe8f52-97db-41f1-b972-4b46e9f1e8d2)

**Video:** [Conditional formatting, charts & dashboards — video tutorials](https://www.youtube.com/results?search_query=excel+conditional+formatting+chart+dashboard+tutorial)

Take the [Stage 5 quiz](../quizzes/module-5) to wrap up. To go further later — each its own topic — look at **Power Query** (auto-refresh from the system), **XLOOKUP & dynamic arrays**, and **Power BI**.
