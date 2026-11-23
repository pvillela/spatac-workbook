# Tracking project timing and cost

Spatac and the Board change when projects happen and what they are expected to cost. Each
change is recorded on **Expenditures**, against the item it affects, together with the reason
for it.

## The columns, left to right

| Columns | Content |
|---|---|
| A–F | The item: row number, item number, building element, GL account, GL account name (folded), and **S** for a Spatac row (folded; one with no RFS amounts, such as a GL account split out of an RFS item or a project the RFS does not foresee). These columns are fixed: unprotect the tab to change them. |
| G–I | Lookup keys (folded) and the [RFS check](analyse-rfs.md#the-rfs-check-column). |
| J–N | **Current year.** Approved (approved by the Board the previous November; grey, fixed for the year), Actuals to date, Mid-year estimate, Late-year estimate, and the calculated **Current estimate**. |
| O–Q | **Next year.** Mid-year estimate, Late-year estimate, **Current estimate**. |
| R–T | **The year after.** The same three. |
| U | **Reason for change** for the next year (headed, for example, "2027 Reason for change"): one short line the Board can read, explaining the approved change from the RFS. |
| V | **Working notes**: the running log of updates and sources. |
| then | The **RFS** block and the **Revised** block, one column per year. |

Above the headings:

- Row 4 totals the columns that hold an amount for every item: Approved, Actuals to date, the
  three Current estimates, and the RFS and Revised blocks. The estimate columns that hold
  changes only have no total.
- Row 5 is the **variance vs RFS** of each Current estimate and each Revised year: its total
  minus the RFS total for the same year. A positive variance means more spending than the RFS.
- Row 6 records where each column's figures came from, such as "to Oct 26" or "2026-11-04".
  Update it when you update a column.

## How the Current estimate is chosen

- **Current year:** the Late-year estimate if one is entered, otherwise the Mid-year estimate,
  otherwise the Approved amount.
- **Next year and the year after:** the Late-year estimate, otherwise the Mid-year estimate,
  otherwise the RFS amount.

Estimates are entered **only for items whose estimate changed**. Type the new full amount, not
the difference. A blank means no change. Type **0** to record that an item will cost nothing
that year, for example because it was deferred.

The two look different: a **0** shows as `0`, and a blank shows nothing. The same holds in the
green Current estimate columns: `0` means the item is now expected to cost nothing that year,
and a blank means it has no amount at all, neither an estimate nor an Approved or RFS amount. In the other amount columns (Approved, Actuals, and the RFS and Revised blocks) a zero
shows as blank, since there it means the same as no amount.

An estimate can be a formula as well as a number, when that is how it was worked out. For
example, `=MAX(J8,K8)` keeps the Late-year estimate at the approved amount until actual spending
exceeds it, then follows actual spending. A cell holding a formula counts as entered, even when
the formula's result is zero.

The **Revised** block takes the three Current estimate columns for the first three years and the
RFS amounts after that. It is what Medium_Term_Cashflow and Projected_Balances use.

## Recording a change

1. Find the item's row. To add a row for something the RFS does not list, see
   [Adding a row](#adding-a-row).
2. Type the new amount in the estimate column for the year affected, such as next year's
   Late-year estimate in November. If a project moves from one year to another, enter **0** in
   the year it leaves and the amount in the year it arrives.
3. Write the reason in column U, in a line the Board can read ("Deferred to 2028: covered by the
   P1 refurbishment"). Details and sources go in Working notes, column V.

The reasons are the record of the Board's changes from the RFS: the Investment Plan's cover
refers readers to them.

## Adding a row

The item number, element, GL account and the other columns A–F are fixed, so adding or
removing an item needs the tab unprotected.

1. Unprotect Expenditures.
2. Insert a row anywhere between the header row and the grey **▼ End** row, for example just
   above ▼ End. The defined names grow to include it (see
   [Inserting rows](../reference/names.md#inserting-rows)). Never delete the header row or the
   ▼ End row.
3. Copy a neighbouring item row into it, then clear its yellow cells: copying brings along the
   formulas in G–I, N, Q, T and the RFS and Revised blocks.
4. Type the item number, element or GL account name, GL account, and **S** in column F if it is
   a Spatac row. The [RFS check](analyse-rfs.md#the-rfs-check-column) in column I flags the
   row if the S is missing or not needed.
5. Type a **Row #** in column A between those of its neighbours, such as 125.5, so that sorting
   by Row # puts it back in its place.
6. Protect the tab again.

## Filtering and sorting

**Filtering** works while the tab is protected. Each heading in row 7 has a filter button: use
it to show, for example, only one GL account or only the items with a reason for change.

**Sorting** needs the tab unprotected: Excel will not sort rows that contain protected cells
on a protected tab.

1. Unprotect Expenditures.
2. Use a heading's filter button, or Data › Sort, to sort by the column you want:
   - **GL account:** column D.
3. To restore the original order, which is item order, sort by **Row #** (column A). Sorting
   by Item (column B) is not reliable: item numbers are text, so 10.1 would come before 2.1.
4. Protect the tab again.

Sorting never moves the ▼ End row: it has no row number, item or GL account, so it
stays at the bottom. The totals and the plan are the same in any order.
