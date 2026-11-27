# The workbook at a glance

## Tabs

| Tab | What it holds |
|---|---|
| **README** | A short guide to the tabs and colours, with a link to this documentation. |
| **Expenditures** | Every RFS component, one row each, plus Spatac's own rows. For each: the current year's approved amount, actual spending, estimates and working notes; and for every year of the RFS, the RFS amount with the Board's adjustments applied, which the plan uses. |
| **Adjustments** | The Board's changes to the RFS, each with its reason: amounts moved between years, added or removed. |
| **Medium_Term_Cashflow** | The Investment Plan: six years of cash in and out, and the investments planned to keep cash at the minimum target. |
| **Schedule_of_Investments** | Investments held and planned, each with its cash flows by year. |
| **Projected_Balances** | The reserve fund balance year by year to the end of the RFS: the RFS's own projection recomputed and checked against RFS_cashflow, and a what-if section to forecast the impact of actuals and Spatac adjustments. |
| **RFS_info**, **RFS_fvalues**, **RFS_cashflow** | The vendor's RFS spreadsheet, pasted in unchanged: the component inventory, the future value of each component's costs by year, and the cash flow summary. |
| **Settings** | The base year (the current year), the minimum target cash balance and the year whose dollars it is stated in, the interest rate assumed for planned investments, the number of years the RFS applies its short-term inflation and interest rates, and the years between RFS updates. |

## Colours

- **Yellow cells are inputs.** They are the only cells you type into.
- **Green cells are calculated** and sit next to inputs, such as the Current estimate columns
  on Expenditures. Read them, don't type into them.
- **Grey cells are fixed for the year**, such as the current year's Approved amounts on
  Expenditures. They change only at the [January rollover](procedures/rollover.md).
- **Orange cells** on Expenditures are adjusted RFS amounts that differ from the RFS because of
  the Adjustments tab.
- Everything else is a label or a calculation.

Every tab is protected **without a password**, the three RFS tabs included, so that a stray
keystroke cannot overwrite a formula or the vendor's figures. Yellow cells stay editable while a
tab is protected. To change anything else, unprotect the tab:

- **Excel:** Review › Unprotect Sheet.
- **LibreOffice:** Tools › Protect Sheet (untick it).

Protect it again when you are done.

Rows with yellow cells have an automatic height: a row grows or shrinks to fit its text as you
type, even on a protected tab, so there is no need to drag a row's border (LibreOffice does not
allow that on a protected tab). At some zoom levels, 110% for one, LibreOffice can draw a fitted
row one line short on screen; the row is right, and shows whole at 100% and when printed.

Some tabs group columns, so that a block of input columns can be folded away, for example when
sharing your screen on a call. The **+** and **−** buttons above the column letters fold and
unfold a group. Excel does not allow this on a protected tab, so unprotect the tab first.

## Defined names

Formulas refer to other tabs through **defined names**, such as `Cfg_BaseYear` for the base
year on Settings, or `Exp_Adjusted` for the block of adjusted RFS amounts on Expenditures. A name keeps
pointing at the right cells when rows or columns are inserted. The command-line tools find
everything through these names too, so do not rename or delete them. The
[reference](reference/names.md) lists them all.

## Years

**Settings › Base year** is the current year. Year headings are calculated from it, and so is
every lookup into the RFS, except on Medium_Term_Cashflow: there each year column has its own
typed year, so that the planned investments typed under it keep their year. Changing it in January rolls the workbook forward
a year: see [January rollover](procedures/rollover.md).

## Estimates are entered only where they change

On Expenditures, the Mid-year and Late-year estimate columns hold an entry only for
items whose estimate changed:

- Type the **new full amount**, never a difference. It can be a formula, such as
  `=MAX(J8,K8)`.
- A **blank** cell means "no change": the item keeps its previous estimate.
- A **0** means the item is now expected to cost nothing that year, for example because it was
  deferred. It shows as `0`, so it cannot be mistaken for a blank.

The green **Current estimate** column beside them shows the resulting amount for every item.
