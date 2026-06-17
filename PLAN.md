# Monthly Close Variance Analysis Skill + Workbook Generator

## Summary

Build a small finance tool plus supporting Codex context files for monthly charter-school close analysis. The v1 tool reads two Excel files for one school: an EdTec GL detail export and the school financial model. It produces an Excel variance workbook that helps a finance analyst quickly identify which GL accounts need review.

Primary comparison: YTD actual activity from the GL detail export vs. `Current Forecast` from the budget/financial model.

The tool is not making final financial conclusions. It performs auditable calculations and flags review items for analyst follow-up.

## Key Changes / Interfaces

### Inputs

- One GL detail export, like `EdTec-GLDetailReport-byAccount,Resource600.xls`.
- One financial model workbook, like `Francophone - FY26 February Financials draft 20260318.xlsm`.
- Analyst chooses where to save the generated output workbook.

### Matching and Calculations

- Match actuals to forecast primarily by `Account Number`.
- Roll detailed GL benefit accounts to the visible forecast parent account when the parent exists, such as `3303` and `3304` rolling to `3300`.
- Calculate YTD actuals from GL detail as net debit/credit by account.
- Use visible account rows with data on the model's `YTD` tab as the forecast source.
- Use model sections to classify account type, such as revenue, expense, or capex.
- Detect close month from model text such as `As of Feb FY2026`.
- Use a month-based forecast usage benchmark: February = 8/12, March = 9/12, etc.

### Output Workbook

- `Summary` tab: separate Revenue Accounts and Expense Accounts tables with account number/name, YTD actual, current forecast, remaining dollars, percent used, benchmark variance, transaction count, latest activity date, and stacked YTD Actual vs. Current Forecast column charts. The revenue chart should include only under-benchmark accounts, the expense chart should include only over-benchmark accounts, each chart should label bar segments with spaced abbreviated dollar values, and wrapped account names should appear below each chart in bar order.
- `Detail` tab: full imported GL detail for analyst drill-down.
- `Problems` tab: actual GL accounts not listed on a visible YTD forecast row, file/sheet/column issues, hidden-row scope notes, and total actuals tie-out status.
- Actual-only revenue and expense accounts from the GL file should also appear on the `Summary` tab. Balance sheet accounts such as `9xxx` remain excluded.

## Build Plan

1. Create the planning/context files, starting with `PLAN.md` and `context/business-rules.md`.
2. Build a local workbook generator that accepts the two source files and asks where to save the output.
3. Parse workbook structure:
   - Read GL detail headers: account, type, date, debit, credit, balance, description, memo, account number, resource, function, site.
   - Read financial model tabs, especially `YTD`.
4. Normalize data:
   - Aggregate GL actuals by account number.
   - Pull current forecast by account number from visible YTD account rows with data.
   - Carry forward model section labels for account type.
5. Calculate review fields:
   - YTD actual.
   - Current forecast.
   - Remaining dollars.
   - Percent forecast used.
   - Favorable/unfavorable direction.
   - Month-based benchmark.
6. Flag accounts:
   - Use remaining dollars plus percent-used logic.
   - Default sensitivity: review accounts meaningfully over the month-based benchmark and material by roughly `$5k and 10%`.
   - Flag unmatched accounts separately.
7. Generate the output workbook with `Summary`, `Detail`, and `Problems`.
8. Add simple stacked column charts showing revenue and expense YTD actuals compared to current forecast.

## Test Plan

- Test with the provided sample files first.
- Confirm output workbook is created successfully.
- Confirm GL actuals aggregate by account number.
- Confirm forecast values are pulled from visible `YTD` rows in the model's `Current Forecast` column.
- Confirm close month is detected from the model title.
- Confirm total actuals tie-out appears in `Problems`.
- Confirm unmatched accounts are listed, not hidden.
- Confirm the demo can show source files -> run tool -> flagged workbook.

## Assumptions

- v1 supports one school at a time.
- The primary user is a finance analyst.
- `Current Forecast` is full-year forecast, so the tool reports percent of full-year forecast used rather than pretending it is a pure YTD budget variance.
- The first success measure is time saved: reducing a 1-2 hour manual review into a repeatable workbook-generation step plus analyst review.
- A floater has reviewed and approved this plan for implementation.
