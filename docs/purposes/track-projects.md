# Tracking project timing and cost

Spatac and the Board change when projects happen and what they are expected to cost. Each
change is recorded on **Expenditures**, against the item it affects, together with the reason
for it.

## The columns, left to right

| Columns | Content |
|---|---|
| A–F | The item: row number (folded), item number, building element, GL account, GL account name (folded), and **S** for a Spatac row (one with no RFS item, such as a GL account split out of an item). |
| G–I | Lookup keys (folded) and the [RFS check](analyse-rfs.md#the-rfs-check-column). |
| J–N | **Current year.** Approved (approved by the Board the previous November), Actual to date, Mid-year estimate, Latest estimate, and the calculated **Current estimate**. |
| O–Q | **Next year.** Mid-year estimate, Late-year estimate, **Current estimate**. |
| R–T | **The year after.** The same three. |
| U | **Reason for change**: one short line the Board can read. |
| V | **Working notes**: the running log of updates and sources. |
| then | The **RFS** block and the **Revised** block, one column per year. |

Row 5 above the headings records where each column's figures came from, such as "to Oct 26" or
"2026-11-04". Update it when you update a column.

## How the Current estimate is chosen

- **Current year:** the Latest estimate if one is entered, otherwise the Mid-year estimate,
  otherwise the Approved amount.
- **Next year and the year after:** the Late-year estimate, otherwise the Mid-year estimate,
  otherwise the RFS amount.

Estimates are entered **only for items whose estimate changed**. Type the new full amount, not
the difference. A blank means no change. Type **0** to record that an item will cost nothing
that year, for example because it was deferred.

An estimate can be a formula as well as a number, when that is how it was worked out. For
example, `=MAX(J7,K7)` keeps the Latest estimate at the approved amount until actual spending
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

Every item whose Revised amounts differ from the RFS in the next six years appears in the
Investment Plan's **Changes from the RFS** list with its reason. A missing reason shows as
REASON MISSING.

## Adding a row

1. Unprotect Expenditures.
2. Insert a row between two existing item rows, so that the defined names grow to include it.
3. Copy the row above, then clear its yellow cells: copying brings along the formulas in G–I,
   N, Q, T and the RFS and Revised blocks.
4. Type the item number, element or GL account name, GL account, and **S** in column F if it is
   a Spatac row.
5. Protect the tab again.
