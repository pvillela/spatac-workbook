# Excel checklist

The restructured workbook is built and checked in LibreOffice. These checks confirm that Excel
sees the same thing. Run them:

- after each restructure;
- after any change to how a tab is built;
- whenever the workbook has only been saved from LibreOffice for a while.

## Opening

- [ ] The file opens without a repair prompt.
- [ ] All nine tabs are present, starting with README and ending with RFS_cashflow.
- [ ] Formulas › Name Manager lists the defined names (about 120), none showing `#REF!`.

## Values

Compare with the same cells in LibreOffice, or with the figures below for the November 2026
workbook:

- [ ] **Medium_Term_Cashflow row 21** (closing cash): 1,039,813 · 1,061,208 · 1,082,190 ·
      1,104,081 · 1,126,162 · 1,148,686.
- [ ] **Medium_Term_Cashflow row 10** (held maturities, 2026): 655,399.
- [ ] **Schedule_of_Investments rows 49 onward:** the planned investments, starting "Planned
      2-year investment, bought 2026". The list must be filled, not blank: it relies on
      `AGGREGATE`.
- [ ] **Expenditures row 4:** totals in the RFS block equal row 3.
- [ ] **Projected_Balances G4:I10:** the lowest margins show amounts and years, not errors.

## Behaviour

- [ ] Typing in a yellow cell works while the tab is protected. Typing in any other cell is
      refused.
- [ ] Review › Unprotect Sheet works without a password.
- [ ] After unprotecting Expenditures, the + and − buttons fold and unfold the column groups.
- [ ] Changing Settings › Base year to 2027 moves every year heading, and Expenditures' RFS block
      shows 2027 first. Undo afterwards.

## Saving

- [ ] Save as `.xlsx` under a new name, then run `xlsx-diff` from the LibreOffice-saved version
      to the Excel-saved one. Without `--formats` it should report no cell changes.
