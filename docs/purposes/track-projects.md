# Tracking project timing and cost

Spatac and the Board change when projects happen and what they are expected to cost. Two tabs
record it:

- **Expenditures** holds the current year's figures for every item (approved amounts, actuals and
  estimates) and shows the RFS amounts for every year with the Board's changes applied.
- **Adjustments** records each change to the RFS, with its reason.

## Expenditures, left to right

| Columns | Content |
|---|---|
| A–F | The item: row number, item number, building element, GL account, GL account name (folded), and **S** for a Spatac row (folded; one with no RFS amounts, such as a GL account split out of an RFS item or a project the RFS does not foresee). These columns are fixed: unprotect the tab to change them. |
| G–I | Lookup keys (folded) and the [RFS check](analyse-rfs.md#the-rfs-check-column). |
| J | **Adjustment label** (folded): the name the Adjustments tab knows the item by (see [Adjustment labels](#adjustment-labels)). |
| L–P | **Current year.** Approved (approved by the Board the previous November; grey, fixed for the year), Actuals to date, Mid-year estimate, Late-year estimate, and the calculated **Current estimate**. |
| Q | **Working notes**: the running log of updates and sources for the current year. |
| S | **Next year**: the adjusted RFS amount for next year, from the block to its right. |
| U onwards | The **adjusted RFS** block: one column per year of the RFS's 30-year span (2024–2053 for the 2024 study), the RFS amount with the Adjustments tab applied. Years before the current year are folded. |

Columns K, R and T are narrow spacers. Working notes name the columns they refer to by their
headings, in quotes, such as "Actuals to date", so they stay right when columns move.

Above the headings:

- Row 3 is the **RFS total (sum of items)** for each year: the amounts of every item on
  RFS_fvalues added up, for next year and each year of the adjusted block. **Years where items
  differ from RFS_fvalues totals**, to the left of the current year's total, counts the years
  where that sum differs from the vendor's own annual total by more than 0.5; those years show
  in red with bold white text. It should be 0.
- Row 4 totals the columns that hold an amount for every item: Approved, Actuals to date, the
  Current estimate, next year and the adjusted block. The estimate columns, which hold changes
  only, have no total.
- Row 5 is the **variance vs RFS** of the Current estimate, next year and each adjusted year:
  the RFS total for the same year minus its total. As on Medium_Term_Cashflow, a **negative
  variance is unfavourable**: more spending than the RFS, shown in red.
- Row 6 records where each current-year column's figures came from, such as "to Oct 26".
  Update it when you update a column. Over the adjusted block, row 6 holds the years the
  formulas read: green and invisible. Never change it; the headings in row 7 show the years.

## The current year

The **Current estimate** is the Late-year estimate if one is entered, otherwise the Mid-year
estimate, otherwise the Approved amount.

Estimates are entered **only for items whose estimate changed**. Type the new full amount, not
the difference. A blank means no change. Type **0** to record that an item will cost nothing
this year.

The two look different: a **0** shows as `0`, and a blank shows nothing. The same holds in the
green Current estimate column: `0` means the item is now expected to cost nothing this year, and
a blank means it has no amount at all, neither an estimate nor an Approved amount. Approved and
Actuals also show a zero as `0`.

Next year and the adjusted block follow the same rule as the Current estimate: an amount shows
blank when the RFS has none for that year and no adjustment touches it, and `0` when the
adjustments leave nothing, for example when a project's whole amount moved to another year.

An estimate can be a formula as well as a number, when that is how it was worked out. For
example, `=MAX(L8,M8)` keeps the Late-year estimate at the approved amount until actual spending
exceeds it, then follows actual spending. A cell holding a formula counts as entered, even when
the formula's result is zero.

## What the plan uses

Medium_Term_Cashflow and Projected_Balances take **the current year from the Current
estimate** column and **every later year from the adjusted RFS block**. The adjusted block's own
column for the current year shows the RFS with the adjustments applied, for comparison; it is
not used for the current year.

An adjusted amount that differs from the RFS amount by more than 50 cents is shaded orange.

## The Adjustments tab

Each row of Adjustments is one change to the RFS:

| Column | Content |
|---|---|
| A | **Item**, picked from a dropdown of the [adjustment labels](#adjustment-labels). |
| B | **Year from**: the year the amount leaves, or **none** when the amount is new money. |
| C | **Year to**: the year the amount arrives, or **none** when the amount is removed. |
| D | **Amount**, greater than 0, typed or given by a formula (see below). |
| E | **Reason for change**: one line the Board can read. |
| F | From-year amount (calculated): what leaves Year from. |
| G | To-year amount (calculated): what arrives in Year to. |
| H | **Check** (calculated): blank when the row is in order. |

Three kinds of row:

- **A move** gives both years. Type the amount in **Year from's dollars**, normally the RFS
  amount being moved. The To-year amount is that amount inflated, or deflated for a move to an
  earlier year, at the RFS's own rates: the short-term rate for the study's first years (the
  number is on Settings) and the long-term rate after them.
- **An addition** has **none** in Year from. Type the amount in **Year to's dollars**.
- **A removal** has **none** in Year to. Type the amount in Year from's dollars.

A part of an amount can be moved; several rows can move the same item. Rows add up: a later row
can move an amount that an earlier row brought in. An amount can be a formula, such as
`=600000-G42` for "the difference between 600,000 and the amount the row above brings in".

The **Check** column shows **OK** for a row that is in order, and is blank for an empty row.
It flags:

- **Missing entry**: a row with some but not all of Item, Year from, Year to, Amount and Reason.
- **Item not in Expenditures**: an Item that is not a label.
- **Year from = Year to**.
- **More than the item holds in *year***: a move or removal larger than the item's RFS amount
  for Year from, after the rows above it.

Cell A2 counts the flagged rows.

### Blocks

The tab has four blocks, one for each of the first four years of the RFS study ("Approved in
2024" to "Approved in 2027" for the 2024 study). Each records the changes the Board approved
in that year. All four apply together.

- At the [January rollover](../procedures/rollover.md), start using the next block.
- With a [new RFS](../procedures/new-rfs.md), all four blocks are emptied, and the first block
  is used again.

### Adding rows to an Adjustments block

Each block has 10 rows. If a block fills up:

1. Unprotect Adjustments.
2. Insert a row inside the block, above its grey closing row.
3. Copy the row above it and paste it into the new row: that brings the dropdowns, the input
   look and the formulas in F, G and H. Then clear the new row's yellow cells.
4. Protect the tab again.
5. Run [`spatac-check`](../tools/check.md): it reports any formula or dropdown the new row is
   missing.

## Adjustment labels

The label in Expenditures column J (folded; unfold it to see the labels) identifies an item on the Adjustments tab: the GL account,
then the building element, or the GL account name when the row has no element. A Spatac row's
label ends in "(Spatac)". Rows without a GL account, such as section headings, have no label;
they show as blank lines in the dropdown, and the Check column rejects them.

## Adding a row to Expenditures

The item number, element, GL account and the other columns A–F are fixed, so adding or
removing an item needs the tab unprotected.

1. Unprotect Expenditures.
2. Insert a row anywhere between the header row and the grey **▼ End** row, for example just
   above ▼ End. The defined names grow to include it (see
   [Inserting rows](../reference/names.md#inserting-rows)). Never delete the header row or the
   ▼ End row.
3. Copy a neighbouring item row into it, then clear its yellow cells: copying brings along the
   formulas in G–J, P, S and the adjusted block.
4. Type the item number, element or GL account name, GL account, and **S** in column F if it is
   a Spatac row. The [RFS check](analyse-rfs.md#the-rfs-check-column) in column I flags the
   row if the S is missing or not needed.
5. Type a **Row #** in column A between those of its neighbours, such as 125.5, so that sorting
   by Row # puts it back in its place.
6. Protect the tab again, and run [`spatac-check`](../tools/check.md).

## Filtering and sorting

**Filtering** works while the tab is protected. Each heading in row 7 has a filter button: use
it to show, for example, only one GL account.

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
