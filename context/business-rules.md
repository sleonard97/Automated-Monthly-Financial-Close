# Business Rules

## Scope

- v1 analyzes one school at a time.
- The primary user is a finance analyst performing monthly financial close review for a charter school.
- The tool compares YTD actual GL activity to the current full-year forecast in the financial model.

## Source Files

- Actuals source: EdTec GL detail export.
- Forecast source: school financial model workbook.
- Forecast value: `Current Forecast` from the model's `Detail` tab.
- Close month: read from model text such as `As of Feb FY2026`.

## Matching

- Primary matching key is `Account Number`.
- Accounts present in actuals but not forecast must be listed on the `Problems` tab.
- Accounts present in forecast but not actuals must be listed on the `Problems` tab.
- Do not silently drop unmatched accounts.

## Actuals Calculation

- Calculate YTD actuals by summing GL activity by account number.
- Use net debit/credit from the GL export as the actuals calculation.
- Keep supporting GL detail available for analyst drill-down.

## Forecast Usage

- `Current Forecast` is treated as a full-year value.
- Do not label the comparison as a pure YTD budget variance.
- Calculate remaining dollars as `Current Forecast - YTD Actual`, with account type and sign convention handled clearly.
- Calculate percent forecast used as `YTD Actual / Current Forecast` when forecast is nonzero.
- Use a month-based benchmark:
  - July = 1/12
  - August = 2/12
  - September = 3/12
  - October = 4/12
  - November = 5/12
  - December = 6/12
  - January = 7/12
  - February = 8/12
  - March = 9/12
  - April = 10/12
  - May = 11/12
  - June = 12/12

## Flagging

- Default sensitivity is roughly `$5k and 10%`.
- Flag accounts that are materially over the month-based usage benchmark.
- Use remaining dollars plus percent-used logic for full-year forecast comparisons.
- Separate data problems from financial review flags.

## Favorable / Unfavorable

- Use model sections to classify account type, such as revenue, expense, or capex.
- Revenue under expectation is generally unfavorable.
- Expense over expectation is generally unfavorable.
- Expense under expectation or revenue over expectation may still require review, but should not be described as inherently bad.

## Validation

- Include a total actuals tie-out check where possible.
- If a tie-out cannot be performed, explain why on the `Problems` tab.
- The tool should create an issue report instead of failing silently when inputs are incomplete or malformed.

