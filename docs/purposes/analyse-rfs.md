# Analysing the Reserve Fund Study

The vendor delivers the RFS as a report (PDF) and as a spreadsheet without formulas. The
spreadsheet's three tabs are pasted, unchanged, into the workbook's RFS tabs.

| Tab | What the vendor's figures are |
|---|---|
| **RFS_info** | The component inventory: item number, GL account, description, history and recommendations, normal expected life, years remaining, and the estimated cost in study-year dollars. |
| **RFS_fvalues** | The repair and replacement schedule: for each item, the cost in each year from the study year on, inflated to that year (its *future value*). Rows 10–13 hold the yearly totals, interest, contributions and balance. |
| **RFS_cashflow** | The cash flow summary: for each year, expected costs, interest, contributions and the closing balance, plus the study's parameters (opening balance, minimum balance, interest and inflation rates, contribution increases). |

Columns to the right of the vendor's data are added by the workbook (RFS_info J:K, RFS_fvalues
BC). They are not part of the paste.

## Where the RFS figures appear

**Expenditures** has one row per RFS item, in the order of the study. Its **RFS** block (the
columns headed by years after the Reason and Working notes columns) shows each item's RFS
amount for every year from the base year on. Each amount is looked up in RFS_fvalues by the
item's number and description, so the block always matches the pasted study.

Row 3 of the RFS block is the vendor's own total for each year and row 4 adds up the items. The
two should agree; a difference means an item is missing from Expenditures. Row 5 shows, for
each year, how far the Revised amounts depart from the RFS.

## Questions the workbook answers

- **What does the RFS expect to spend in a given year, and on what?** Filter or sort the RFS
  block's year column on Expenditures.
- **Which items drive a heavy year?** For example 2033, with slab and garage roof waterproofing
  and the main roof. Sort that year's column, largest first.
- **How does Spatac's view differ from the RFS?** The Revised block, to the right of the RFS
  block, holds the amounts the plan uses. The two blocks share their column order, so
  comparing a year in one with the same year in the other shows the difference.
  Medium_Term_Cashflow rows 28–30 compare the yearly totals for the next six years.
- **What did the study assume?** RFS_cashflow rows 6–12 hold its parameters. Projected_Balances
  shows what they produce year by year; see the [next chapter](validate-funding.md).

## The RFS check column

Expenditures column **I** (RFS check) flags a row the RFS no longer supports:

- **Not in RFS**: the row's item number and GL account are not in RFS_info.
- **No RFS amounts**: the item is in RFS_info, but its number and description were not found in
  RFS_fvalues, so its RFS amounts show as zero.

Spatac's own rows, marked **S** in column F, are not checked. The flags matter mostly after a
new RFS has been pasted in: see [A new Reserve Fund Study](../procedures/new-rfs.md).
