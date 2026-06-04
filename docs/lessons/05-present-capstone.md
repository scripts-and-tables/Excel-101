---
layout: lesson
title: "Stage 5 — Present & capstone"
short_title: "Present & capstone"
subtitle: "Make the numbers land on one page — then run the whole workflow yourself."
module: 5
---

> **Where we are:** Receive & clean → Structure & enrich → Explore & summarize → Pivot & rank → **Present**.

You've done the analysis. The last step of the workflow is turning it into something the VP can read in ten seconds: a clean one-pager. Open `module-5.xlsx` and build it on the `One-Pager` sheet.

## 1. A KPI strip

Put the headline numbers across the top, big and formatted — they're just the Stage 3 KPIs:

- Total Closed-Won revenue · Order count · Average Order Value · Win rate · Refund rate

Currency as `$1.2M`, rates as `%`. Top-left is where eyes land — put the most important number there.

## 2. One clean chart

You rarely need more than one. Build a `Region | Revenue` summary, select it, **Insert → Column Chart**. Then: rename the title to something specific, delete gridlines and the legend (single series), format the axis as currency, and **sort the source descending** so the bars rank themselves.

## 3. Let conditional formatting do the flagging

On a summary table or the data itself (**Home → Conditional Formatting**):

- **Color scales** turn a Region × Quarter grid into a heat map.
- **Data bars** rank reps in-cell — perfect on the leaderboard.
- **Icon sets** add traffic lights (e.g. on quota attainment).
- **Highlight → Duplicate Values** flags repeated customers/IDs — a fast second check on Stage 1's dedupe.
- **Formula rule:** New Rule → *Use a formula* → `=$L2="Refunded"` → red fill, and the **whole row** of every refunded order turns red.

Manage it all from **Conditional Formatting → Manage Rules**.

## 4. Layout tips

- **Top-left = most important.** Headline KPI goes there.
- **Group related things**, keep **one colour per region** everywhere, and never show a raw `1234567.89` — use `$1.2M` or `$1,234,567`.
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
| `Ctrl + Shift + 1 / 4 / 5` | Number / Currency / Percent |
| `Alt + =` | AutoSum |
| `Alt + N + V` | Insert PivotTable |
| `F2` | Edit the active cell |
| `F4` | Toggle absolute/relative reference |

---

## 🎯 Capstone — run the whole workflow

Open `capstone.xlsx` (working). It's a **fresh sales file** plus the `Reps` table — the same shape you started with, new data. Your job is the entire workflow, end to end.

> **Brief:** You're the new RevOps analyst. Before Monday's meeting, answer the VP's questions on the `Questions` sheet (column C), then drop the headline numbers into a one-page summary. Use PivotTables, AutoFilter, `SUMIFS`/`COUNTIFS`, and `VLOOKUP` — whatever's fastest.

The 11 questions (on the `Questions` sheet):

1. Total revenue from **Closed Won** orders only.
2. Total order count.
3. **Refund rate** (Refunded ÷ all).
4. **Win rate** (Closed Won ÷ all).
5. **Average Order Value** on Closed Won.
6. Closed Won revenue for the **West** region.
7. Closed Won revenue for the **Hardware** category.
8. Closed Won revenue booked by **Anna Becker**.
9. Anna Becker's **quota attainment** (her Closed-Won revenue ÷ her `AnnualQuota` — a `VLOOKUP`).
10. Anna Becker's **commission** (her Closed-Won revenue × her `CommissionRate` — a `VLOOKUP`).
11. Average **discount** given on Closed Won orders.

**Ground rules:** only count `Status = "Closed Won"` where asked; format money as `$1,234,567` and rates/discounts as `%`; prefer a **live formula** over a typed number so it survives a data refresh.

Compare against `capstone.xlsx` (solution) — the `Answers` sheet has a working formula for each. There's often more than one correct route.

## You finished the course

You can now take a raw sales export and walk it through the whole process: **clean → enrich → explore & summarize → pivot & rank → present**. That's the 80% of Excel a sales team actually uses, in the order you'd actually use it.

Take the [Stage 5 quiz](../quizzes/module-5) to wrap up. To go further later — each its own topic, out of scope here — look at **Power Query** (auto-refresh from the CRM), **XLOOKUP & dynamic arrays**, and **Power BI** (when a workbook outgrows Excel).
