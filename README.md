# Automated Monthly Financial Close

Generate an Excel workbook that helps a finance analyst review monthly close
variance items for one charter school.

## Run the workbook generator

Install the dependency:

```powershell
python -m pip install -r requirements.txt
```

Run the generator:

```powershell
python scripts/generate_close_variance.py `
  --gl "C:\path\to\EdTec-GLDetailReport-byAccount,Resource600.xls" `
  --model "C:\path\to\Francophone - FY26 February Financials draft 20260318.xlsm" `
  --output "C:\path\to\Monthly Close Variance Review.xlsx"
```

If you leave off any file path, the script will ask for it.

## Output

The generated workbook includes:

- `Summary`: separate Revenue Accounts and Expense Accounts tables with variance fields, plus stacked column charts comparing YTD Actual to Current Forecast. The revenue chart includes only accounts under the benchmark; the expense chart includes only accounts over the benchmark. Charts show abbreviated dollar labels spaced across each bar segment, with wrapped account-name labels below each chart in bar order.
- `Detail`: imported GL detail for drill-down.
- `Problems`: actual GL accounts not listed on a visible YTD forecast row, input issues, hidden-row scope notes, and the total actuals tie-out check.

The report uses only visible account rows with data on the model's `YTD` tab.
Hidden account rows in the forecast workbook are excluded from the analysis.
Each `Summary` table follows the same relative account order as the visible `YTD` tab.
Revenue and expense accounts with GL activity but no visible YTD forecast row
are appended to the relevant `Summary` table for review. Balance sheet accounts such as
`9xxx` are excluded from that actual-only list.
Detailed GL benefit accounts roll up to the visible forecast parent account when
available, such as `3303` and `3304` rolling up to `3300`.
