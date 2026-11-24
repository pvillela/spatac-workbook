# Excel checklist

The restructured workbook is built and checked in LibreOffice. These checks confirm that Excel
sees the same thing. Run them:

- after each restructure;
- after any change to how a tab is built;
- whenever the workbook has only been saved from LibreOffice for a while.

## Opening

- [ ] The file opens without a repair prompt.
- [ ] All ten tabs are present, starting with README, Expenditures and Adjustments and ending
      with Settings.
- [ ] Formulas › Name Manager lists the defined names (about 120), none showing `#REF!`.

## Values

Compare with the same cells in LibreOffice, or with the figures below for the November 2026
workbook:

- [ ] **Medium_Term_Cashflow row 20** (closing cash): 1,039,813 · 1,061,208 · 1,082,190 ·
      1,104,081 · 1,126,162 · 1,148,686.
- [ ] **Medium_Term_Cashflow row 10** (held maturities, 2026): 655,399.
- [ ] **Schedule_of_Investments rows 49–78:** the planned investments, one row per grid cell,
      with gaps; row 50 reads "Planned 2-year investment, bought 2026". No cell shows an error.
- [ ] **Formulas with `LET`, `IFS` and `SWITCH`** calculate (Excel 2021 or later), for example
      Expenditures I8, J8 and T8, Adjustments H7, Medium_Term_Cashflow B9 and B11, and
      Projected_Balances K17 and W17. Their line
      breaks show in the formula bar when it is expanded.
- [ ] **Expenditures row 5** over the adjusted block (from column T): 0 in every year but
      2025 (−188,700), 2026 (−416,160), 2027 (607,552) and 2028 (151,541), the years the
      Adjustments tab changes.
- [ ] **Expenditures, orange shading:** the adjusted amounts that differ from the RFS are
      orange, for example row 29 (3.9) in 2025 and 2028. Unfold the past years to see 2025.
- [ ] **Adjustments, 2026 block:** column H is blank on all six rows; G shows 243,547 · 50,000
      · 432,973 · 514,686 · 85,314 · 40,800.
- [ ] **Projected_Balances G4:I10:** the lowest margins show amounts and years, not errors.
- [ ] **Projected_Balances validation:** choose **RFS** in the dropdowns B4, B6 and B7: column Q
      shows 0 in every year. Undo afterwards.

## Behaviour

- [ ] Typing in a yellow cell works while the tab is protected. Typing in any other cell is
      refused.
- [ ] Review › Unprotect Sheet works without a password.
- [ ] After unprotecting Expenditures, the + and − buttons fold and unfold the column groups.
- [ ] **Expenditures N33** (a Late-year estimate of 0) shows `0`; a blank estimate cell, such
      as L33, shows nothing.
- [ ] **Adjustments dropdowns:** in an empty row of the 2027 block, A offers the item labels,
      B and C offer "none" and 2024–2053. D refuses 0 and text, and accepts `=1000*2`. With
      the same year in B and C, H shows "Year from = Year to". Undo.
- [ ] **Expenditures RFS check:** column I is blank on every row. Unprotect the tab, type **S**
      in F11: I11 shows "In RFS: clear S". Undo, and protect the tab again.
- [ ] On the protected Expenditures tab, the filter buttons in row 7 work: filter column D on
      one GL account, then clear the filter.
- [ ] Unprotect Expenditures and sort by column D (GL account): the grey ▼ End row stays at
      the bottom, and row 4's totals do not change.
      Sort by **Row #** (column A): the original order returns. The cell comments move with
      their rows. Protect again.
- [ ] Insert a row just below the header row (row 7) of Expenditures and another just above
      ▼ End: Formulas › Name Manager shows `Exp_Item` grown by two rows. Undo.
- [ ] The README tab's long lines show in full (wrapped, not cut to one line).
- [ ] On a copy: unprotect Medium_Term_Cashflow, delete column B, insert a column before
      ▶ End and fill it from its left neighbour. Row 11 shows numbers, not errors, and the
      Planned list on Schedule_of_Investments is still filled. Discard the copy.
- [ ] Changing Settings › Base year to 2027 moves the year headings, and Expenditures' next-year
      column (R) is headed "2028 Adjusted". Medium_Term_Cashflow keeps its years and A2 shows the red check. Undo
      afterwards.

## Saving

- [ ] Save as `.xlsx` under a new name, then run `xlsx-diff` from the LibreOffice-saved version
      to the Excel-saved one. Without `--formats` it should report no cell changes.
