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

1. **Approved for the new year.** Copy the next-year **Current estimate** (column Q, from row 8 down to the row above ▼ End)
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

Each year column has its own year, typed in row 3, and keeps its figures when the base year
changes. Rolling forward means removing the year just ended and adding a new last year. Until
you do, the red check in A2 says the first year column is not the base year.

1. **Check the old year's planned investments.** Every planned investment for the year just
   ended (column B, rows 14–18) should have been bought and typed into Schedule_of_Investments
   as held. Its maturities then come from there. Anything left in column B is lost in the next
   step.
2. **Delete column B**, the year just ended. Everything to its right moves one column left, with
   its year. Column B's opening cash (B5) shows `#REF!` until step 5.
3. **Insert a column** just left of the grey **▶ End** column (right-click the ▶ End column
   heading › Insert Columns Before). Never delete the ▶ End column.
4. **Fill it from the column to its left:** copy that column and paste it into the new one.
   Then type the new year in row 3 (one more than its neighbour) and clear its yellow cells
   (payables, the 1- to 4-year investments, the carry-over and the adjustment) or type the
   planned figures.
5. Type the new year's **opening cash balance** (B5) and **payables** (B6) from the prior year's
   closing financial statements.
6. Check that A2 is empty, then update the date in the title (A1).

## Schedule_of_Investments

1. For each investment that matured last year, clear its yellow cells (A–E and G), or delete
   its row.
2. Type each investment bought last year into an empty yellow row. If none is left, insert rows
   just above the grey **▼ End of held investments** row and copy the formulas in F and H–X
   from the row above. Never delete the Held label row or the ▼ End row.
3. Update the "as at" date in the title (A1).

## Projected_Balances

Type the new year's **opening balance** (B5) from the prior year's financial statements: cash,
less payables and accrued liabilities, plus investment principal and accrued interest. The
what-if now starts in the new year.

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
