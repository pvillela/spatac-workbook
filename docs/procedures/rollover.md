# January rollover

Once a year, in mid to late January, the workbook moves forward a year. The base year goes
from, say, 2026 to 2027: 2027 becomes the current year, 2028 the next year, and 2029 the year
after. Do it once the prior year's closing figures are in and the Board's November decisions
for the new year are recorded.

**The order matters.** The Current estimate columns fall back to the RFS for the year they
show. Copy them while they still show the old years, which means before changing the base year.

## Before you start

1. Save the workbook under the new date, such as `Spatac-20270120.xlsx`. The previous file is the
   record of last year.
2. Unprotect Expenditures, Medium_Term_Cashflow, Schedule_of_Investments and Projected_Balances.

## Expenditures

1. **Approved for the new year.** Copy the next-year **Current estimate** (column Q, rows 8–251)
   and paste **values only** into **Approved** (column J).
   - Excel: Paste Special › Values.
   - LibreOffice: Paste Special › Values Only.

   This is the amount the Board approved in November, for every item.
2. **Next year's starting estimate.** Copy the year-after **Current estimate** (column T) and
   paste values only into the next-year **Mid-year estimate** (column O). It stands until the
   next mid-year review replaces it.
3. **Clear** the old year's inputs: Actuals to date (K), Mid-year estimate (L) and Late-year
   estimate (M). Also clear the next-year Late-year estimate (P) and the year-after estimates (R and S).
4. **Reasons.** Review column U: delete reasons for changes that no longer apply.
5. Update the "as of" notes in row 6.

## Settings

Change **Base year** (B4) to the new year. Every year heading and RFS lookup moves with it.

## Medium_Term_Cashflow

1. Type the new year's **opening cash balance** (B5) and **payables** (B6) from the prior year's
   closing financial statements.
2. **Move the planned investments one column left.** Copy C14:G17 and paste values only into
   B14:F17, then clear G14:G17. Do the same for the carry-over (row 22) and the adjustment
   (row 23).
3. Update the date in the title (A1).

## Schedule_of_Investments

1. For each investment that matured last year, clear its yellow cells (A–E and G).
2. Type each investment bought last year into an empty yellow row.
3. Update the "as at" date in the title (A1).

## Projected_Balances

For the year just ended, type the **actual expenditures** (column F) and the **actual closing
balance** (column G) from the financial statements.

## Finish

1. Protect the four tabs again.
2. Compare last year's file with this one:

   ```
   xlsx-diff Spatac-20261121.xlsx Spatac-20270120.xlsx -o rollover.md
   ```

   Read the report. Expect the base year, the pasted columns, the cleared inputs and the
   recalculated totals. Nothing should have changed in a formula.
3. Extract the Investment Plan for approval: see
   [spatac-investment-plan](../tools/investment-plan.md).
