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
4. **Empty the Adjustments tab.** The adjustments changed the old study; the new one starts
   clean. Unprotect Adjustments and clear the yellow cells (columns A–E) of all four blocks.
   Leave the green columns F–H: their formulas stay. The block titles follow the new study year
   by themselves, and changes the Board approves from now on go in the first block.

## 3. Reconcile the item list

The new study should keep the current item numbers and GL accounts, but a few may change. Two
check columns show where:

- **Expenditures column I** flags rows without S whose item and GL account are no longer in
  RFS_info ("Not in RFS"), or whose item and description are not in RFS_fvalues ("No RFS
  amounts"), and rows with S that the study now covers ("In RFS: clear S"). See
  [The RFS check column](../purposes/analyse-rfs.md#the-rfs-check-column).
- **RFS_info column K** flags RFS items with no row on Expenditures.

For each flag:

- **An item was renumbered or redescribed:** update the item number, description or GL account
  on Expenditures to match the study.
- **An item is new:** add a row for it; see
  [Adding a row](../purposes/track-projects.md#adding-a-row-to-expenditures).
- **An item was dropped:** if Spatac still expects to spend on it, mark the row **S** in column
  F. Otherwise clear its inputs, or delete the row.
- **A Spatac row is now in the study** ("In RFS: clear S"): clear the **S** in column F. If the
  row now duplicates an RFS item's own row, merge the two instead.

Column F is protected: unprotect Expenditures to change it, and protect it again afterwards.

When both columns are clear, every RFS item has its row and every row finds its amounts. Then
check, on Expenditures, that row 3 (the RFS totals) equals row 4 (the items added up) in
the adjusted block. (With the Adjustments tab emptied in section 2, the two rows agree in every
year.)

## 4. Review what follows from the study

- **Projected_Balances P6:** the last year of contribution increases fixed by the new funding
  plan. Review the catch-up increases in P7:P9.
- **Projected_Balances, validation:** I13 shows 0: the recomputed RFS balance matches
  RFS_cashflow in every year (see
  [Checking the vendor's arithmetic](../purposes/validate-funding.md#checking-the-vendors-arithmetic)).
- **Settings:** the minimum target and its year, if the Board revisits them; the **years of
  short-term inflation**, which must match the study's (RFS_fvalues labels them, for example
  "Inflation 0-3" for 3); and the **years of short-term interest** (RFS_cashflow F7, for
  example "Short Term (3 yr.) Interest").
- **Expenditures:** review the current year's estimates against the new study.
- **Adjustments:** record in the first block any change from the new study that the Board
  approves; see [The Adjustments tab](../purposes/track-projects.md#the-adjustments-tab).

Compare the files before and after with `xlsx-diff` and read the report.
