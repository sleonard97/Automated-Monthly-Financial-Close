# Workflow

## Current Monthly Close Workflow

1. Finance analyst receives or exports GL detail for a school.
2. Analyst opens the school's financial model.
3. Analyst compares actual activity against budget or forecast information.
4. Analyst filters, sums, and researches account-level variances manually.
5. Analyst identifies accounts that need review, explanation, or follow-up.
6. Analyst prepares notes for close review.

## Target v1 Workflow

1. Analyst provides the GL detail export and the school financial model.
2. Tool reads the files and identifies the close month from the model.
3. Tool aggregates actuals by account number.
4. Tool matches actuals to current forecast by account number.
5. Tool calculates YTD actual, current forecast, remaining dollars, percent used, and month-based benchmark.
6. Tool creates an Excel output workbook.
7. Analyst reviews flagged accounts in the `Summary` tab.
8. Analyst drills into the `Detail` tab when support is needed.
9. Analyst uses the `Problems` tab to resolve unmatched accounts, missing columns, or tie-out concerns.

## Demo Story

Show the source files, run the tool, then open the generated workbook with prioritized flags. The main value story is reducing a 1-2 hour manual review into a repeatable workbook-generation step plus analyst review.

