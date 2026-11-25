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
      Expenditures I8, J8 and U8, Adjustments H7, Medium_Term_Cashflow B9 and B11, and
      Projected_Balances D15, P16, Q16 and AB16. Their line
      breaks show in the formula bar when it is expanded.
- [ ] **Expenditures row 5** over the adjusted block (from column U): 0 in every year but
      2025 (188,700), 2026 (416,160), 2027 (−607,552, red) and 2028 (−151,541, red), the years the
      Adjustments tab changes.
- [ ] **Expenditures, orange shading:** the adjusted amounts that differ from the RFS are
      orange, for example row 29 (3.9) in 2025 and 2028. Unfold the past years to see 2025.
- [ ] **Adjustments, 2026 block:** column H shows OK on all six rows; G shows 243,547 · 50,000
      · 432,973 · 514,686 · 85,314 · 40,800.
- [ ] **Projected_Balances B4:G7 and R4:W7:** the lowest and first negative margins show
      amounts and years, not errors (`MINIFS` needs Excel 2019 or later): T5 −1,882,352 (2050),
      V5 −1,058,672 (2045).
- [ ] **Projected_Balances money format:** a negative amount such as U43 (−341,094) shows red
      with a minus sign.
- [ ] **Projected_Balances validation:** column I shows 0 in every year and I12 shows 0. Type
      1000 more into RFS_cashflow G20: I19 turns orange and I12 shows 1 in red. Undo afterwards.
- [ ] **Projected_Balances, light grey:** rows 44–63 (2054–2073) are shaded.

## Behaviour

- [ ] Typing in a yellow cell works while the tab is protected. Typing in any other cell is
      refused.
- [ ] Review › Unprotect Sheet works without a password.
- [ ] After unprotecting Expenditures, the + and − buttons fold and unfold the column groups.
- [ ] **Expenditures O33** (a Late-year estimate of 0) shows `0`; a blank estimate cell, such
      as M33, shows nothing.
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
      column (S) is headed "2028 Adjusted". Medium_Term_Cashflow keeps its years and A2 shows the red check. Undo
      afterwards.

## Saving

- [ ] Save as `.xlsx` under a new name, then run `xlsx-diff` from the LibreOffice-saved version
      to the Excel-saved one. Without `--formats` it should report no cell changes.
