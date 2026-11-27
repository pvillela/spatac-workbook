# spatac-check

`spatac-check` checks the integrity of a Spatac workbook. Run it after editing the workbook by
hand, especially after inserting rows, and before sharing it.

```
spatac-check Spatac-20261124.xlsx
```

Save the workbook in Excel or LibreOffice first: the tool reads the saved file and does not
recalculate it.

## What it checks

- **Missing formulas.** Every formula column must have a formula in every row: the
  Expenditures label, key, check, Current estimate, next-year and adjusted columns, the
  Adjustments From-year amount, To-year amount and Check columns, and the calculated rows of
  Medium_Term_Cashflow and Projected_Balances. A row inserted into an Adjustments block starts
  without them; see [Adding rows to an Adjustments block](../purposes/track-projects.md#adding-rows-to-an-adjustments-block).
- **Names.** A defined name the workbook relies on that is missing, or that points at `#REF!`
  because the cells it named were deleted.
- **Errors.** A cell whose last calculated result is an error, such as `#N/A` or `#REF!`.
- **Validations.** An Adjustments input row whose Item, Year from or Year to has lost its
  dropdown, or whose Amount has lost its rule.
- **Protection.** A tab left unprotected (the three RFS tabs included), a formula cell that is
  unlocked, or an Adjustments input cell that is locked.

## The report

Every problem is listed, by tab and cell, in the order of the tabs and rows:

```
Spatac-20261124.xlsx: 3 problem(s)
- Adjustments!F45: a formula is missing (copy it from the row above)
- Adjustments!G45: a formula is missing (copy it from the row above)
- Adjustments!H45: a formula is missing (copy it from the row above)
```

When nothing is wrong it says `no problems found`. The exit status is 0 when the workbook is
clean, 1 when a problem was found, and 2 when the file could not be read.
