# Data Shape

## GL Detail Export

Observed file example: `EdTec-GLDetailReport-byAccount,Resource600.xls`.

Expected relevant columns:

- `Account`
- `Type`
- `Date`
- `JE# - Inv.# - Ck# -`
- `Vendor/Cust. Name`
- `Debit`
- `Credit`
- `Balance`
- `Description`
- `Memo`
- `Account Number`
- `Resource-Yr (Dept)`
- `Function-Goal (Class)`
- `Site (Loc)`

The `.xls` export may be XML-style Excel rather than a binary legacy workbook.

## Financial Model Workbook

Observed file example: `Francophone - FY26 February Financials draft 20260318.xlsm`.

Observed tabs:

- `Validation`
- `YTD`
- `Detail`
- `Cash Flow`
- `Balance Sheet`
- `Graphs`
- `CapEx`

Important tabs for v1:

- `Detail`: current forecast detail by account.
- `YTD`: model title, close month, YTD actuals, approved budget YTD, current forecast, notes.

Observed useful model labels:

- `As of Feb FY2026`
- `Current Forecast`
- `Actual YTD`
- `Approved Budget v1 YTD`
- `YTD Variance`
- `Notes`

## Output Workbook

Required tabs:

- `Summary`
- `Detail`
- `Problems`

