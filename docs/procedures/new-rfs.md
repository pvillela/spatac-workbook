# A new Reserve Fund Study

A new RFS or RFS update arrives every three years. The vendor provides a spreadsheet without
formulas, as well as the report. The spreadsheet replaces the contents of the three RFS tabs.

## 1. Check the layout

The workbook reads the vendor's figures from fixed places, through defined names:

| Tab | Expected layout |
|---|---|
| RFS_info | Item number in A, GL account in B, description in C, cost in H; one row per item. |
| RFS_fvalues | Years in row 9 from column C; yearly totals in row 10; items from row 15, with the item number in A and the description in B. |
| RFS_cashflow | Years in column A from row 15; costs in B, interest in C, contributions in D, balance in G. Parameters: units C6, minimum balance C10, opening balance C11, present contribution C12, interest G7:G8, inflation K7:K8, contribution increases K9:K12. |

If the vendor's spreadsheet is laid out differently, adjust the names under the **Formulas ›
Name Manager** (Excel) or **Sheet › Named Ranges and Expressions › Manage** (LibreOffice)
before pasting. The [reference](../reference/names.md) lists every `RFS_` name.

## 2. Paste

1. Save the workbook under a new date.
2. On each RFS tab, select the vendor's area and delete its contents:
   - RFS_info: A:H;
   - RFS_fvalues: A:BA;
   - RFS_cashflow: A:K.

   Leave the helper columns to the right (RFS_info J:K, RFS_fvalues BC) alone.
3. Paste the vendor's corresponding tab at A1, as values.

## 3. Reconcile the item list

The new study should keep the current item numbers and GL accounts, but a few may change. Two
check columns show where:

- **Expenditures column I** flags rows whose item and GL account are no longer in RFS_info
  ("Not in RFS"), or whose item and description are not in RFS_fvalues ("No RFS amounts").
- **RFS_info column K** flags RFS items with no row on Expenditures.

For each flag:

- **An item was renumbered or redescribed:** update the item number, description or GL account
  on Expenditures to match the study.
- **An item is new:** add a row for it; see
  [Adding a row](../purposes/track-projects.md#adding-a-row).
- **An item was dropped:** if Spatac still expects to spend on it, mark the row **S** in column
  F. Otherwise clear its inputs, or delete the row.

When both columns are clear, every RFS item has its row and every row finds its amounts. Then
check, on Expenditures, that row 3 (the vendor's totals) equals row 4 (the items added up) in
the RFS block.

## 4. Review what follows from the study

- **Projected_Balances B8:** the last year of contributions fixed by the new funding plan.
  Review the catch-up increases in B9:B11.
- **Projected_Balances, validation:** set the what-if to RFS and check that the difference
  from the RFS balance is zero in every year (see
  [Re-doing the vendor's arithmetic](../purposes/validate-funding.md#re-doing-the-vendors-arithmetic)).
- **Settings:** the minimum target and its year, if the Board revisits them.
- **Expenditures:** for the current year and the next two, review the estimates that now differ
  from the new RFS. Clear any that no longer reflect a Spatac decision, so that the RFS amount
  shows through.

Compare the files before and after with `xlsx-diff` and read the report.
