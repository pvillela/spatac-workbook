# January rollover

Once a year, in mid to late January, the workbook moves forward a year. The base year goes
from, say, 2026 to 2027: 2027 becomes the current year and 2028 the next year. Do it once the
prior year's closing figures are in and the Board's November decisions for the new year are
recorded on the Adjustments tab.

**The order matters.** Expenditures' next-year column shows the year after the base year.
Copy it while it still shows the new year, which means before changing the base year.

## Before you start

1. Save the workbook under the new date, such as `Spatac-20270120.xlsx`. The previous file is the
   record of last year.
2. Unprotect Expenditures, Medium_Term_Cashflow, Schedule_of_Investments and Projected_Balances.

## Expenditures

1. **Approved for the new year.** Copy the **next year** column (S, from row 8 down to the row
   above ▼ End) and paste **values only** into **Approved** (column L).
   - Excel: Paste Special › Values.
   - LibreOffice: Paste Special › Values Only.

   This is the RFS amount with the Board's adjustments: what the Board approved in November,
   for every item.
2. **Clear** the old year's inputs: Actuals to date (M), Mid-year estimate (N) and Late-year
   estimate (O).
3. **Working notes** (Q): clear or shorten the notes that were about the year just ended.
4. Update the "as of" notes in row 6.

## Adjustments

Changes approved from now on go in the **next block** ("Approved in 2027" after the 2027
rollover). The earlier blocks stay as they are: all blocks apply together. The adjustments are
not moved or cleared at the rollover; only a [new RFS](new-rfs.md) empties the blocks.

## Settings

Change **Base year** (B4) to the new year. Every year heading and RFS lookup moves with it.

Then, on Expenditures, **fold the year just ended** in the adjusted block: select its column
(for example "Adjusted 2026"), then Data › Group and Outline › Group (Excel: Data › Group), so
that it joins the folded past years, and collapse the group with its **−** button. Past years
do not affect the plan; folding only keeps them out of the way.

## Medium_Term_Cashflow

Each year column has its own year, typed in row 3, and keeps its figures when the base year
changes. Rolling forward means removing the year just ended and adding a new last year. Until
you do, the red check in A2 says the first year column is not the base year.

1. **Check the old year's planned investments.** Every planned investment for the year just
   ended (column B, rows 14–18) should have been bought and typed into Schedule_of_Investments
   as held. Its maturities then come from there. Anything left in column B is lost in the next
   step.
2. **Delete column B**, the year just ended. Everything to its right moves one column left, with
   its year. Column B's opening cash (B5) shows `#REF!` until step 5, and so does everything
   that follows from it: the 5-year investments (row 18), the planned new investments total
   (row 13) and the closing cash (row 20) of every year, and on Schedule_of_Investments the
   Planned 5-year rows. This is expected; it clears in step 5.
3. **Insert a column** just left of the grey **▶ End** column (right-click the ▶ End column
   heading › Insert Columns Before). Never delete the ▶ End column. Each Planned row on
   Schedule_of_Investments shows its own grid cell, so the list now shows the years in their
   new columns.
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

Type the new year's **opening balance** (P5) from the prior year's financial statements: cash,
less payables and accrued liabilities, plus investment principal and accrued interest. The
what-if now starts in the new year.

The November 2026 workbook was delivered with P5 seeded from the old simulator's projected
closing balance for 2025, worked out from the 2024 audited closing balance and the 2025 actual
expenditures. Replace it
with the figure from the 2025 financial statements once they are available.

## Finish

1. Protect the four tabs again.
2. Compare last year's file with this one:

   ```
   xlsx-diff Spatac-20261121.xlsx Spatac-20270120.xlsx -o rollover.md
   ```

   Read the report. Expect the base year, the pasted columns, the cleared inputs and the
   recalculated totals. Nothing should have changed in a formula.
3. Run [`spatac-check`](../tools/check.md) on the new file: it should find no problems.
4. Extract the Investment Plan for approval: see
   [spatac-investment-plan](../tools/investment-plan.md).
