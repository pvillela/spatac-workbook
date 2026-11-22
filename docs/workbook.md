# The workbook at a glance

## Tabs

| Tab | What it holds |
|---|---|
| **README** | A short guide to the tabs and colours, with a link to this documentation. |
| **Expenditures** | Every RFS component, one row each, plus Spatac's own rows. For each: the amount approved, actual spending, estimates and the reason for any change; the RFS amount for every year; and the Revised amount the plan uses. |
| **Medium_Term_Cashflow** | The Investment Plan: six years of cash in and out, and the investments planned to keep cash at the minimum target. |
| **Schedule_of_Investments** | Investments held and planned, each with its cash flows by year. |
| **Projected_Balances** | The reserve fund balance year by year to the end of the RFS, with Spatac's changes and what-if contribution increases. |
| **RFS_info**, **RFS_fvalues**, **RFS_cashflow** | The vendor's RFS spreadsheet, pasted in unchanged: the component inventory, the future value of each component's costs by year, and the cash flow summary. |
| **Settings** | The base year (the current year), the minimum target cash balance and the year whose dollars it is stated in, and the interest rate assumed for planned investments. |

## Colours

- **Yellow cells are inputs.** They are the only cells you type into.
- **Green cells are calculated** and sit next to inputs, such as the Current estimate columns
  on Expenditures. Read them, don't type into them.
- **Grey cells are fixed for the year**, such as the current year's Approved amounts on
  Expenditures. They change only at the [January rollover](procedures/rollover.md).
- Everything else is a label or a calculation.

The Spatac tabs are protected **without a password**, so that a stray keystroke cannot
overwrite a formula. Yellow cells stay editable while a tab is protected. To change anything
else, unprotect the tab:

- **Excel:** Review › Unprotect Sheet.
- **LibreOffice:** Tools › Protect Sheet (untick it).

Protect it again when you are done.

Some tabs group columns, so that a block of input columns can be folded away, for example when
sharing your screen on a call. The **+** and **−** buttons above the column letters fold and
unfold a group. Excel does not allow this on a protected tab, so unprotect the tab first.

## Defined names

Formulas refer to other tabs through **defined names**, such as `Cfg_BaseYear` for the base
year on Settings, or `Exp_Revised` for the block of Revised amounts on Expenditures. A name keeps
pointing at the right cells when rows or columns are inserted. The command-line tools find
everything through these names too, so do not rename or delete them. The
[reference](reference/names.md) lists them all.

## Years

**Settings › Base year** is the current year. Every year heading in the workbook is calculated
from it, and so is every lookup into the RFS. Changing it in January rolls the workbook forward
a year: see [January rollover](procedures/rollover.md).

## Estimates are entered only where they change

On Expenditures, the Mid-year and Late-year estimate columns hold an entry only for
items whose estimate changed:

- Type the **new full amount**, never a difference. It can be a formula, such as
  `=MAX(J8,K8)`.
- A **blank** cell means "no change": the item keeps its previous estimate.
- A **0** means the item is now expected to cost nothing that year, for example because it was
  deferred.

The green **Current estimate** column beside them shows the resulting amount for every item.
