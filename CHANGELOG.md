# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## Unreleased

### Added

- `spatac-check`, a third tool: checks a Spatac workbook's integrity and lists every problem
  (missing formulas, broken or missing names, cells showing errors, Adjustments inputs without
  their validation, protection that does not match a cell's role). Exit status 1 on a problem.
- The workbook's **Adjustments** tab (design 14): the Board's changes to the RFS, each a move
  between years, an addition or a removal, with its reason, in four blocks by approval year.
  Amounts move at the RFS's own short-term and long-term inflation rates.

### Changed

- The workbook, 2026-11-27 (design 17): the RFS tabs are protected and `spatac-check` requires
  it; the RFS totals check sits in Expenditures T1 with its alarm; the Projected_Balances alarms
  show white text (their own number format) and column I is rounded to whole dollars; every
  input row has an automatic height; Expenditures headings keep "Mid-year" and "Late-year"
  whole, column C wraps, and columns too narrow for their text on six tabs are widened.
- `tools/migrate/`: `layout_lint.py` checks what a build looks like on every tab (fit, word
  breaks, colour overrides, orphaned formats, "-0", row heights, protection), and `render.py`
  prints tabs to PDF to look at; every build runs both before it is handed over.
- Expenditures: the next-year and year-after estimate blocks, the Reason for change column and
  the RFS and Revised blocks are replaced by an adjustment label (J), a next-year column (R) and
  the adjusted RFS for the study's 30 years (T onwards), shaded where it differs from the RFS.
  Working notes moved next to the current year, with their column letters updated.
- Medium_Term_Cashflow: a blank row before the minimum, now "Minimum closing cash balance"; the
  RFS row reads RFS_fvalues.
- Projected_Balances: the inflator applies the short-term rate for Settings' new "years of
  short-term inflation" and the long-term rate after.
- `spatac-investment-plan`: the plan keeps its source's fonts, drops the input look of the
  source's input cells, gives the Cover's approval lines the input look, and opens with only
  the Cover selected.
- `spatac-investment-plan`: the Changes_from_RFS tab is gone. The cover says instead that the
  reasons for changes from the RFS are described in the Spatac workbooks for the current and
  previous years.
- `spatac-investment-plan`: the cover's "Generated" date is today in the computer's own time
  zone, not UTC; `--date YYYY-MM-DD` sets it.
- `xlsx-diff`: neighbouring cells changed the same way across a row, or over a block, are one
  entry, as runs down a column already were.
- The workbook restructure (`tools/migrate/`), from the Board's review of 2026-11-21: top
  alignment and grid borders; Settings last; Expenditures with protected identity columns, a
  fixed Approved column, a variance row, restored comments, filtering, sorting and ▼ End rows;
  Projected_Balances rebuilt as the RFS baseline and a what-if that validates RFS_cashflow.
- Expenditures: a zero in the estimate columns L–T shows as `0`, no longer as an empty cell,
  so that "will cost nothing" (0) and "no change" (blank) look different. The Current
  estimate columns N, Q and T are empty, not 0, when the item has no estimate, Approved or RFS
  amount.
- Workbook formulas simplified (design 13.10): `LET`, `IFS` and `SWITCH` replace nested IFs
  and long expressions; the Expenditures item sort key is gone (sort by Row #); the Planned
  investments list has one fixed row per grid cell; the Projected_Balances what-if has Opening
  balance and Interest rate columns, so the difference from the RFS balance is now column Q.
- `tools/migrate/migrate.py` always applies the phase 4 fixes; its `--fix` options and
  `verify.py --expect-changes` are gone.
- Expenditures: the RFS check also flags a Spatac row (S) that the RFS covers, "In RFS: clear
  S", so a stale S no longer hides the row from the check.
- Expenditures: Row # and Item (columns A and B) are left-aligned, as in the original sheet, and the title is in A1.

### Fixed

- `xlsx-diff`: a column whose formulas use absolute references to columns that a column insert
  had moved (`$U8:$AX8` becoming `$V8:$AY8`) was reported as deleted and re-inserted, with every
  one of its cells as a change, and a defined name on it as a range whose ends crossed
  (`$U$7:$R$250`). Columns are now aligned again with the old formulas rewritten through the
  first alignment, and a range whose rows or columns were all deleted reads `#REF!`.
- `xlsx-diff`: a cell unlocked through its cell style was read as locked when its own format
  set `applyProtection="false"`, as LibreOffice writes it, so saving a workbook in another
  application reported cells as newly unlocked. The unlocked state is now read from the file's
  formats.
- The workbook build kept no frozen panes, and lost comments on formula cells.

## 0.1.0 - 2026-11-21

### Added

- `xlsx-diff`: compares two `.xlsx` workbooks and writes the differences as Markdown.
  - Tabs added, deleted, renamed (recognised by content) and reordered; defined names.
  - Per tab: rows and columns inserted, deleted and moved; changed formulas, constants and
    comments. Formulas are compared after allowing for moved rows and columns, and changes
    filled down a column are grouped into one entry.
  - Structural changes: merged cells, conditional formats, data validation, protection,
    unlocked cells, hidden rows, columns and tabs, frozen panes.
  - `--values` for recalculated results, `--formats` for formats, widths and heights.
- `spatac-investment-plan`: extracts the Investment Plan from a Spatac workbook into
  `InvestmentPlan-YYYYMMDD.xlsx`, with:
  - a cover;
  - Medium_Term_Cashflow and Schedule_of_Investments, where formulas referring to other tabs
    become values;
  - a list of changes from the RFS with their reasons.

  Every cell is locked and every tab protected without a password.
- `--third-party-notices` on both tools; release builds embed notices generated by
  `scripts/gen-notices.sh`.
- Documentation (mdBook, published to GitHub Pages): the workbook's five purposes, the January
  rollover and new-RFS procedures, editing in Excel or LibreOffice, an Excel checklist, the tools,
  and a reference of defined names.
- `tools/migrate/`: the LibreOffice scripts that restructured the workbook, verified it against
  the original, and check an extracted Investment Plan.
