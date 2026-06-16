# Monthly Close Variance Analysis Skill + Workbook Generator

## Summary

Build a small finance tool plus supporting Codex context files for monthly charter-school close analysis. The v1 tool reads two Excel files for one school: an EdTec GL detail export and the school financial model. It produces an Excel variance workbook that helps a finance analyst quickly identify which GL accounts need review.

Primary comparison: YTD actual activity from the GL detail export vs. `Current Forecast` from the budget/financial model.

The tool is not making final financial conclusions. It performs auditable calculations, flags review items, and drafts short analyst-style review prompts.

## Key Changes / Interfaces

### Inputs

- One GL detail export, like `EdTec-GLDetailReport-byAccount,Resource600.xls`.
- One financial model workbook, like `Francophone - FY26 February Financials draft 20260318.xlsm`.
- Analyst chooses where to save the generated output workbook.

### Matching and Calculations

- Match actuals to forecast primarily by `Account Number`.
- Calculate YTD actuals from GL detail as net debit/credit by account.
- Use the model's `Detail` tab `Current Forecast` as the forecast source.
- Use model sections to classify account type, such as revenue, expense, or capex.
- Detect close month from model text such as `As of Feb FY2026`.
- Use a month-based forecast usage benchmark: February = 8/12, March = 9/12, etc.

### Output Workbook

- `Summary` tab: flagged GL accounts with account number/name, type, YTD actual, current forecast, remaining dollars, percent used, flag reason, favorable/unfavorable label, and short AI review prompt.
- `Detail` tab: full imported GL detail for analyst drill-down.
- `Problems` tab: unmatched actual accounts, unmatched forecast accounts, file/sheet/column issues, and total actuals tie-out status.

### AI Behavior

- AI writes concise analyst notes only.
- AI uses calculated fields, account labels, and flags.
- AI must not quote or use vendor/customer names, descriptions, or memos.
- AI notes should be review prompts, not cause claims or action directives.

## Build Plan

1. Create the planning/context files, starting with `PLAN.md` and `context/business-rules.md`.
2. Build a local workbook generator that accepts the two source files and asks where to save the output.
3. Parse workbook structure:
   - Read GL detail headers: account, type, date, debit, credit, balance, description, memo, account number, resource, function, site.
   - Read financial model tabs, especially `Detail` and `YTD`.
4. Normalize data:
   - Aggregate GL actuals by account number.
   - Pull current forecast by account number from the model.
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
8. Add AI-generated review prompts to the `Summary` tab.
9. Optional polish after core workflow works: add simple charts/KPI visuals showing top flagged accounts and forecast usage.

## Test Plan

- Test with the provided sample files first.
- Confirm output workbook is created successfully.
- Confirm GL actuals aggregate by account number.
- Confirm forecast values are pulled from the model's `Current Forecast`.
- Confirm close month is detected from the model title.
- Confirm total actuals tie-out appears in `Problems`.
- Confirm unmatched accounts are listed, not hidden.
- Confirm AI notes do not quote vendor names, descriptions, or memos.
- Confirm the demo can show source files -> run tool -> flagged workbook.

## Assumptions

- v1 supports one school at a time.
- The primary user is a finance analyst.
- `Current Forecast` is full-year forecast, so the tool reports percent of full-year forecast used rather than pretending it is a pure YTD budget variance.
- The first success measure is time saved: reducing a 1-2 hour manual review into a repeatable workbook-generation step plus analyst review.
- A floater has reviewed and approved this plan for implementation.

