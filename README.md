# BudgetWise Personal Finance — Excel Analytics Project

An end-to-end Excel analytics project built on a messy, synthetic personal-finance transaction dataset. The focus is on demonstrating a realistic data-cleaning pipeline (Power Query), analytical formula work, PivotTable reporting, and light automation (VBA) — the kind of workflow used in audit and financial-analysis settings.

## Project Overview

| | |
|---|---|
| **Dataset** | BudgetWise Personal Finance Dataset (synthetic, ~15,000 transaction records) |
| **Tools** | Microsoft Excel — Power Query, Excel Tables, PivotTables, formulas, VBA |
| **Focus** | ETL, data-quality auditing, spend analysis, reporting automation |

The dataset was intentionally messy — inconsistent date formats, mixed currency symbols, misspelled/abbreviated categorical values, placeholder strings standing in for missing data, and a handful of invalid transaction amounts. The project's goal was to clean, standardize, and analyze it using an auditable, fully reproducible Excel workflow — no manual one-off edits, everything driven by Power Query steps and live formulas.

## 1. Data Cleaning (Power Query / ETL)

All cleaning logic lives in the `BudgetTransaction` query and re-runs from the original source CSV on every refresh — nothing here is a manual, one-time fix.

**Key transformations:**
- **Date standardization** — custom M logic to parse multiple inconsistent date formats (slash-delimited, dash-delimited, 2-digit vs. 4-digit years) into a single clean `date` type column
- **Currency & numeric cleanup** — stripped mixed currency symbols (`Rs.`, `INR`, `$`, `₹`) and non-numeric characters from the `amount` field, converting it to a clean numeric type
- **Categorical standardization** — mapped inconsistent/misspelled entries in `category`, `payment_mode`, and `location` (e.g., "fod", "foodd" → `Food`) to a controlled set of values using `List.Contains` logic
- **Placeholder-to-null normalization** — values like `"No notes provided"`, `"N/A"`, and `"Unmapped"` were standardizing text, not real nulls. These were explicitly replaced with true `null`s so that completeness metrics and filters treat them consistently with genuinely missing data
- **Invalid row filtering** — removed transactions with negative amounts (142 rows), which don't represent valid spend/income entries in this dataset
- **Data Quality Flag** — added a `data_quality_flag` column (`Complete` / `Missing Data`) based on whether `date` or `amount` is null, enabling a completeness KPI without fabricating any values

**A deliberate decision:** no financial figures were imputed anywhere in this pipeline. Missing dates and amounts are left null and flagged rather than estimated — mirroring standard audit practice, where guessing a transaction amount or date is a bigger risk than simply flagging it as incomplete.

## 2. Analysis Layer (Excel Formulas)

Built on the `Summary_Analysis` sheet, using the cleaned Excel Table as the source (`SUMIFS`, `COUNTIF`, `COUNTA`):

- **Data Completeness KPI** — Total Transactions, Complete Records, Missing Data Records, and % Complete, computed directly off the `data_quality_flag` column
- **User-Level Spend Summary** — Total Expense, Total Income, and Net position calculated per unique `user_id` using `SUMIFS`, spilled across all ~150 users

## 3. Reporting (PivotTable)

A `Budget_vs_Actual` PivotTable breaks down Expense totals by **category** (rows) and **month** (columns), giving a category × time view of spend patterns. A separate `Budget` table holds average monthly actuals against budget targets per category, for a budget-vs-actual comparison.

## 4. Automation (VBA)

A `RefreshEverything` macro, triggered by a button on the `Summary_Analysis` sheet, refreshes the full workflow in one click:
1. Triggers `RefreshAll` on the workbook's data connections and PivotTables
2. Waits for all query connections to fully finish refreshing (avoids a timing issue where formulas could recalculate against partially-refreshed data)
3. Forces a full recalculation of every formula in the workbook
4. Confirms completion with a message box

This means the entire cleaning-through-reporting pipeline can be re-run end-to-end from a single button, against an updated source CSV, with no manual steps in between.

## Key Takeaways

- Built a fully reproducible ETL pipeline in Power Query — no manual data edits, everything reruns cleanly from the raw source file
- Applied a deliberate, documented approach to missing data: flag, don't fabricate
- Combined formula-based and PivotTable-based reporting to demonstrate both approaches to the same underlying analysis
- Automated the refresh-and-recalculate workflow with VBA for one-click reproducibility

## Files

- `BudgetTransaction_Clean.xlsm` — full workbook (Power Query, formulas, PivotTable, and VBA macro)
- `budgetwise_finance_dataset.csv` — original raw source data

---
