# Spatac workbook — design decisions

Decision record from the design session of 2026-11-21. It is the specification the
restructured workbook and the command-line tools are built from. It is kept for future
reference and is not end-user documentation.

Question numbers (Qn) refer to that session and are kept so a decision can be traced
back to its discussion.

## 1. Context

The Spatac workbook supports the Antiquary condominium (MTCC No. 1360) board and SPATAC
(Special Projects and Technical Advisory Committee) in five purposes:

1. Analyse the Reserve Fund Study (RFS), which the vendor supplies as a formula-free
   spreadsheet.
2. Validate the RFS funding proposal (rate and horizon assumptions).
3. Track Spatac adjustments to project timing and cost.
4. Support investment planning. Under the Reserve Fund Investment Policy (s.3.2) the
   Medium-Term Cash Flow schedule *is* the Investment Plan: a 5-year horizon, matching
   maturities plus cash to anticipated expenditures, and identifying Board changes to
   the RFS with their reasons.
5. Estimate the contribution increases needed after the next RFS.

The overarching goal is user-friendliness.

## 2. Users and roles (Q2)

- SPATAC and Board members (e.g. Bruce) use the workbook. Workbook documentation is
  written for non-technical readers.
- The CLI tools are run by the workbook maintainer. Windows and Linux binaries are
  released anyway.
- The InvestmentPlan workbook is a polished, read-only document for the Board, the
  Treasurer and the broker.
- This may be revisited later.

## 3. Applications and spreadsheet features

### 3.1 Applications (Q5, Q18)

- The workbook must work in current Excel and in LibreOffice Calc 26.2, and be editable
  in either.
- The master format is `.xlsx`, named `Spatac-YYYYMMDD.xlsx`.
- Edits made in parallel in another copy (e.g. Bruce in Excel) are found with the diff
  tool against the version he started from, then re-applied by hand to the master
  (Q18 b). There is no patch/merge tool.

### 3.2 Test results, LibreOffice 26.2.7.3 (Q17, Q24)

Measured in this devcontainer:

- There is no dynamic-array spilling. Entering an array-returning formula normally creates
  a legacy array formula sized to the result at entry time. It never resizes: growth is
  ignored and shrinkage gives `#N/A`. A blocked spill gives no `#SPILL!`.
- `A1#` spill references are not supported (Err:508).
- `LAMBDA` is not supported (Err:509).
- `LET`, `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, `XLOOKUP`, `VSTACK`/`HSTACK`,
  `CHOOSECOLS`, `TAKE`/`DROP` and `TOCOL` exist and evaluate. They save with the correct
  `_xlfn.` prefixes.
- LibreOffice saves array formulas as legacy CSE arrays without dynamic-array metadata.
  An Excel-style dynamic array loads as a fixed array, and its metadata is dropped on
  re-save.
- `AGGREGATE(15;6;…/(cond);k)` works as an ordinary single-cell formula and survives a
  round trip through `.xlsx`.

### 3.3 Formula rules (Q4, Q17)

- **No spill formulas** anywhere: no `#` references, no `LAMBDA`.
- Do not enter formulas as spills in Excel. The docs tell Excel users how to avoid this.
- **Default style:** identical single-cell formulas filled across a block, reading defined
  names, e.g. `SUMPRODUCT`, `INDEX`/`MATCH`, `XLOOKUP`, `AGGREGATE`. This replaces the
  original goal of "new-style array formulas".
- `LET`, `FILTER`, `XLOOKUP` etc. are allowed **inside** formulas that return a single
  value.
- Legacy fixed arrays only for blocks whose shape never changes, and preferably not at
  all.
- Forbidden: `OFFSET`, `INDIRECT`, VBA/macros, Power Query, Excel Tables and structured
  references. `Table1` is removed (Q4 i).
- **Readability (13.10):** a formula should be understandable from the formula bar. Use `LET`
  to name the parts of a complex formula, `IFS` or `SWITCH` instead of nested `IF`s, a layout
  over several lines for long formulas, and helper columns where they make the logic easier to
  follow. `IFS` and `SWITCH` need Excel 2019 or later; `LET`, Excel 2021 or later.

### 3.4 Naming (Q30)

- Defined names are workbook-scoped, with a per-tab prefix plus PascalCase, e.g.
  `MTC_ClosingCash`, `Inv_Due`, `Exp_RFS`, `RFS_Fv_Costs`, `Cfg_BaseYear`.
- A defined name may cover a 2-D block (rows × year columns) with a companion header
  name, e.g. `Exp_RFS` and `Exp_RFSYears`.
- CLI tools refer only to defined names, never to cell addresses.

### 3.5 Inputs and protection (Q6)

- An input cell may hold a formula as well as a number: a formula is how a spreadsheet records
  how an amount was worked out. "Blank" in the Current estimate rules means an empty cell.
  (An earlier draft said estimate cells hold no formulas; the user overruled it after phase 7.)
- Input cells use a named `Input` cell style and are unlocked. Computed columns that sit
  next to inputs use a distinct `Calculated` style.
- Sheets are protected **without a password**, to prevent accidental edits only.
- Excel's outline (+/−) buttons do not work on a protected sheet, so collapsing or expanding the
  column groups (Q46) means unprotecting the sheet first. Decided during phase 3: protect every
  Spatac tab anyway.
- The README tab carries the legend.

## 4. Workbook structure

### 4.1 Tabs (Q14, Q41)

| Tab | Source | Notes |
|---|---|---|
| README | new | Purpose of each tab, input legend, link to the documentation site (Q12) |
| Settings | new | `Cfg_BaseYear` (single cell, drives the rollover), minimum target amount and its base year, assumed new-investment rate |
| Expenditures | `fvalues_spatac` | Renamed (Q32); see 4.2 |
| Medium_Term_Cashflow | same | See 4.3 |
| Schedule_of_Investments | same | See 4.4 |
| Projected_Balances | `cashflow_spatac` + `cashflow_spatac_simulator` | Merged (Q14 c); see 4.5 |
| RFS_info, RFS_fvalues, RFS_cashflow | `info`, `fvalues`, `cashflow` | Untouched paste-in zones for vendor data, plus helper and check columns outside the paste area |

Removed:

- `fvalues_spatac_end_2025` (Q14 b). The dated workbook file is the snapshot.
- `P1 Project` (Q41).

### 4.2 Expenditures (Q28, Q29, Q31–Q33, Q35–Q38, Q45, Q46)

**Wording.** The word "budget" is not used anywhere in the workbook.

**Column layout, left to right:**

1. **Identity.** Row #, item, element, GL account, GL account name, `Spatac row` flag.
   - Item, element and GL account stay visible (frozen).
   - GL account name, and other secondary columns, sit in a collapsible column group.
2. **Current year.**
   - Inputs: Approved (full column), Actual to date (full column), Mid-year estimate
     (changes only), Latest estimate (changes only).
   - **Current estimate** (calculated) = Latest estimate, else Mid-year estimate, else
     Approved.
   - No actual-to-date floor (Q39).
3. **Next year.**
   - Inputs: Mid-year estimate and Late-year estimate (changes only).
   - **Current estimate** = Late-year estimate, else Mid-year estimate, else RFS.
4. **Year after.** Same as next year.
5. **Reason for change.** Short and Board-facing (Q33).
6. **Working notes.** The former column M log.
7. **RFS block** (`Exp_RFS`, `Exp_RFSYears`).
   - One column per year, from `Cfg_BaseYear` to the RFS end year.
   - Each cell is a lookup into `RFS_fvalues`, keyed on **item number + description**
     (Q31). The vendor reuses item numbers 4.6 and 18.3.1 for different components.
   - A row whose key is not found is flagged, never silently zero.
8. **Revised block** (`Exp_Revised`).
   - One column per year.
   - The first three years come from the Current estimate columns; later years from the
     RFS block.

**Behaviour:**

- Input columns marked "changes only" are blank wherever there is no change. Their
  input-only columns sit in collapsible groups (Q46), so a meeting can view only the
  Current estimate columns.
- Amounts display as whole dollars, as today.
- Row identity is **item number + GL account**. The GL code is compared as text with 4
  decimals, because `RFS_info` stores it numerically, e.g. `9220.101`.
- **New-RFS checks (Q38).**
  - One column flags Expenditures rows whose item + GL account is absent from
    `RFS_info`. Spatac rows are exempt. (Superseded by 13.4: the check also flags a Spatac
    row the RFS covers.)
  - Another flags `RFS_info` items with no Expenditures row.
- The item list is typed, not formula-generated, so inputs stay attached to their rows
  (Q35).
- Green conditional-format row highlighting is removed.
- The estimate formulas of the old sheet are carried over with their references rewritten:
  column L's `=MAX(I,J)` / `=J` become `=MAX(J,K)` / `=K` in the Latest estimate; N30, O70 and
  Q70 look up the item's RFS amount for the year they named (2025 or 2026) instead of the
  pasted RFS columns; O208 keeps its constant expression. Results are preserved.
- Usability on laptop screens shared over Zoom is to be judged by the user after the
  build (Q45). If it fails, the fallback is a separate "Meeting view" tab.

### 4.3 Medium_Term_Cashflow (Q20–Q22, Q25–Q27)

**Horizon.** Six year columns, `Cfg_BaseYear` to `Cfg_BaseYear`+5. The headers come from
Settings.

**Revised rows.**

- Rows 9 and 29 are both labelled **Revised**.
- Each is the year-column sum of `Exp_Revised`.
- Row 28 is the RFS equivalent, from `Exp_RFS`.

**Maturities and interest.** Year totals of the Schedule_of_Investments year columns, for
Held + Planned. Nothing is placed by hand.

**Planned new investments grid** (term 1–5 × purchase year):

- It remains the input.
- The 5-year row remains the balancing figure.
- A single assumed rate is used (Q26).
- The hand-built diagonal maturity row (old row 41) is removed.

**Minimum target cash balance (Q22).**

- Inputs on Settings: amount (today $1,000,000) and base year (2024).
- Inflated at the RFS short-term inflation rate. There is no override cell.

**Removed.** Scratch area J37:K41 (Q15, Q27).

### 4.4 Schedule_of_Investments (Q20, Q21, Q25)

**Held section.**

- Typed investments: name, due date, purchase date, amount, rate, compounding flag.
- Room for about 40 rows; blank rows are ignored.

**Planned section.**

- 30 formula rows, filled automatically from the Medium_Term_Cashflow grid.
- They are compact: only the non-zero grid cells, in purchase-year order, found with
  `AGGREGATE`/`INDEX`. Unused rows are blank.

**Year columns** run from `Cfg_BaseYear` to +10.

- Each row shows its interest in each year and its principal + interest in the maturity
  year.
- Maturities after the horizon stay visible here (Q27).

**Workflow.**

- Buying an investment: add a Held row and zero its grid cell.
- At rollover: delete the matured Held rows and add the new ones.

**Removed (Q15).** Columns L:R, the cash diary in rows 37–46, and all `#NAME?` cells.

### 4.5 Projected_Balances (Q14 c, Q42)

- One row per year, 2024–2073 (the RFS range).
- RFS figures are read directly from `RFS_cashflow` (no pasted values).
- Columns, in order:
  - Spatac cost adjustment: Revised minus RFS, every year, not only 2026–2028.
  - Adjusted costs.
  - Contributions: the RFS schedule, or typed increase rates for the catch-up and
    subsequent years.
  - Interest at the RFS rate on the adjusted balance.
  - Adjusted balance.
  - Variance.
  - Minimum-balance margins, PV and inflated.
- Lowest margin over 30 years and its year are displayed. Goal Seek is documented for
  purpose 5.
- A collapsed check column recomputes the vendor's own balance and flags any mismatch
  (purpose 2).

## 5. Clean-up and bugs

**Garbage removed (Q15):**

- 3 invisible broken charts.
- Schedule_of_Investments L:R and rows 37–46.
- Medium_Term_Cashflow J37:K41.
- `fvalues` X8:AZ8 stray numbers.
- Hidden zero rows 217–218.
- Column formatting out to column 1024 and about 8k styled empty cells.
- Orphan cell styles and unused conditional-format styles.
- The label-only "Total adjustments" row.
- Legacy threaded-comment stubs, reviewed rather than blindly deleted, since their text
  is the comment.

**Bugs (Q16 b).** First reproduce today's numbers exactly, then fix each bug in a separate
labelled step:

- Schedule_of_Investments compounding is hard-coded to `^5` regardless of term.
- Medium_Term_Cashflow K41 multiplies by J13, the rate cell. It is removed with the
  scratch area.
- Expenditures totals sum rows 8–252 while the data ends at 250. (Phase 3 found this was not a
  bug: rows 251–252 hold Spatac rows, and only `Table1`'s range was short. The new totals cover
  every row.)
- Medium_Term_Cashflow closing cash balance counts the adjustment (row 22) twice: the 5-year
  balancing figure subtracts it and the closing balance adds it again. Found during phase 3.
  Fix, decided by the user: closing balance `=SUM(B5:B12)`, so that it equals the minimum plus
  the carry-over plus the adjustment, each counted once. The carry-over row moves below the
  closing balance, after a blank row, and the adjustment row is relabelled "Adjustment to
  minimum cash balance".

The minimum balance (Q16, item 5) turned out not to be a bug. It is the Board's $1M
target in 2024 dollars; see 4.3.

## 6. Procedures (documented, not automated; Q13)

### 6.1 January rollover (Q34, Q37)

1. Set `Cfg_BaseYear` to the new year. Headers, the RFS block and the Medium_Term_Cashflow
   years all move.
2. Paste values: the old next-year **Current estimate** becomes the new **Approved**.
3. Carry the old year-after estimate into the new next year. Paste values from the old
   year-after **Current estimate** into the new next-year Mid-year column. That column
   always has a value, because it falls back from the Late-year estimate to the Mid-year
   estimate to the RFS. The old year-after Mid-year entries are not carried separately.
4. Clear Actual to date, Mid-year estimate and Latest estimate.
5. The new year-after columns start empty and fall back to the RFS.
6. Schedule_of_Investments: delete the matured Held rows and add the new ones.

### 6.2 New RFS (Q35, Q38)

1. If the vendor's item/account structure is unchanged, replace the contents of
   `RFS_info`, `RFS_fvalues` and `RFS_cashflow` wholesale.
2. Read the two structure-check columns.
3. Reconcile any flagged rows: add, remove or renumber Expenditures rows.
4. Years before the base year drop out automatically.

### 6.3 Parallel edits from another person (Q18)

1. Diff their copy against the version they started from.
2. Re-apply the changes to the master.

## 7. Restructure method (Q3, Q40)

- A scripted migration (Python driving headless LibreOffice 26.2) reads
  `Spatac-20261121.xlsx` and writes the restructured workbook. The script is one-off dev
  tooling kept in the repo.
- Verification:
  - Key results (Medium_Term_Cashflow totals, Projected_Balances tab) match today's values
    exactly.
  - The diff tool's report is reviewed.
  - The maintainer runs an Excel checklist on each milestone.
- Any change made by hand afterwards is recorded.
- LibreOffice 26.2.7.3 is installed in the devcontainer from the official `.deb`
  packages. Add it to the devcontainer setup permanently.

## 8. Command-line tools (Rust, umya-spreadsheet; Linux and Windows)

### 8.1 Crate (Q8, Q19)

- One crate, two binaries: `xlsx-diff` (generic, no Spatac knowledge) and
  `spatac-investment-plan`.

### 8.2 xlsx-diff (Q9–Q11, Q43)

**Tabs.** New, deleted and renamed tabs are identified only. Rename detection uses content
similarity. Reordering is described.

**Structural changes:**

- Row and column alignment (LCS over row/column signatures) reports inserts, deletes and
  moves.
- Reported by default:
  - defined names;
  - merged cells;
  - conditional formatting;
  - data validation;
  - protection and unlocked cells;
  - hidden rows, columns and tabs;
  - frozen panes.

**Cell changes:**

- Compared on the aligned grid. Formulas are compared in relative (R1C1) form, so shifted
  formulas are not reported as changed.
- Default: formula changes, constant changes and comment changes.
- `--values` adds computed-value changes in a separate collapsed section.
- `--formats` adds styles, widths and heights.
- Identical filled formulas are reported as one change.

**Output.** Markdown, with per-sheet cell lists in `<details>` blocks. There is no JSON
output.

### 8.3 spatac-investment-plan (Q23, Q44)

**Input and output.**

- Input `Spatac-YYYYMMDD.xlsx`; output `InvestmentPlan-YYYYMMDD.xlsx`.
- Locates content only through defined names.

**Kept:**

- Medium_Term_Cashflow cash section and the RFS-vs-Revised comparison.
- Schedule_of_Investments Held and Planned sections with their year columns.
- A new **Changes from the RFS** list: the item, RFS year and amount, revised year and
  amount, and Reason for change. Changed items with a blank reason are flagged.

**Dropped:** the simulator comparison (old rows 32–34) and working notes.

**Formulas.**

- References within the two tabs stay as formulas.
- References to other tabs become values, taken from cached values.

**Presentation.**

- All cells locked; sheets protected without a password.
- Cover block: source file, date generated, base year, and "approved at meeting ___".

### 8.4 Tests (Q7)

- The repo is public. Real workbooks are never committed (they live in `_nogit`).
- Fixtures are generated by Rust test code:
  - small synthetic workbooks for the diff tool;
  - a synthetic workbook with Spatac's tabs and defined names for the extract tool.

## 9. Licensing, CI and release (Q1)

- `MIT OR Apache-2.0`, with `LICENSE-MIT` and `LICENSE-APACHE` (already present).
- Modelled on `pvillela/ev-cost-recovery`:
  - **`ci.yml`:** on every push and PR, `ubuntu-26.04`. Runs clippy with `-D warnings`,
    `cargo test --all-targets --all-features`, and `cargo doc` with intra-doc link
    lints.
  - **`release-build.yaml`:** on `v*` tags.
    - Checks that the tag matches the Cargo.toml version; a hyphen marks a pre-release.
    - Generates third-party notices with cargo-about (`scripts/gen-notices.sh`, gated by
      `build.rs`).
    - Builds Linux (`ubuntu-26.04`) and Windows (`windows-latest`); tests run on the
      Windows leg.
    - Packages `.tar.gz` and `.zip` with the licences and notices, containing **both**
      binaries.
    - Publishes with `softprops/action-gh-release@v2`.
  - A hand-written `CHANGELOG.md` (Keep a Changelog, SemVer).

## 10. Documentation (Q12, Q47)

- Markdown sources built with mdBook and published on GitHub Pages, with chapters for the
  five purposes and for the rollover, new-RFS and parallel-edit procedures.
- The in-workbook README tab is the entry point for board members.
- A PDF export of the Investment Plan chapter is the "written description of the schedule"
  that policy s.3.2 requires.
- `docs/archive/` (this file) is excluded from the published book.

## 11. Deferred and open items

- Usability of the Expenditures width on Zoom screen-share (Q45, Q46): judged by the user
  after the build.
- Estimate columns (Mid-year, Latest, Late-year) hold **replacement amounts, not deltas**.
  They are entered only for items whose estimate changed; blank means no change. The
  Current estimate column shows the resulting amount for every row (Q45, confirmed).
- Excel-side checks still to run: that LibreOffice-saved legacy arrays and the `AGGREGATE`
  pattern behave in Excel.
- Possible later work:
  - a patch/merge tool for parallel edits;
  - a CLI report of RFS structure changes;
  - `import-rfs` and `rollover` automation;
  - per-term assumed rates;
  - a computed minimum contribution increase.

## 12. Phases of work

Work happens on the `aicode` branch. Each phase ends with its verification passing and one
commit; a phase may contain several commits where that aids traceability (e.g. one per bug
fix in phase 4).

Once this document is approved, the phases run in sequence without stopping for approval.
Three things are exceptions:

- **A failed verification is a stop.** If a phase's verification cannot be made to pass
  without departing from this document, work stops and the problem is reported.
- **A design change is a stop.** If implementation shows a decision here is wrong or
  incomplete, work stops and the question is put to the user. This document is updated
  with the answer before continuing.
- **No outward-facing actions.** Nothing is pushed, no tag is created and no release is
  published without the user's explicit instruction.

Real workbooks stay in `_nogit` and are never committed. Generated workbooks are written
there too, as `_nogit/Spatac-YYYYMMDD-restructured.xlsx` and similar.

### Phase 1 — Repository scaffolding

- Commit this design record.
- Cargo layout:
  - one crate with binaries `xlsx-diff` and `spatac-investment-plan`;
  - a library module shared by both;
  - `src/bin/` stubs.
- CI and release modelled on `ev-cost-recovery` (section 9):
  - `ci.yml`;
  - `release-build.yaml`, packaging both binaries;
  - cargo-about (`about.toml`, `about.md.hbs`, `scripts/gen-notices.sh`, `build.rs`
    gate);
  - `CHANGELOG.md`.
- Devcontainer: install LibreOffice 26.2 and its runtime libraries in
  `.devcontainer/setup.sh`.
- README skeleton.

**Verification:** `cargo clippy --all-targets -- -D warnings`, `cargo test` and `cargo doc`
pass locally. The workflow YAML is syntactically valid.

### Phase 2 — `xlsx-diff`

Implement section 8.2 in order:

1. Workbook model loading (umya-spreadsheet).
2. Tab add, delete, rename and reorder.
3. Row and column alignment.
4. Cell diff with relative-formula comparison and grouping of filled formulas.
5. Comments.
6. Structural items (names, merges, conditional formatting, validation, protection,
   hidden items, panes).
7. `--values` and `--formats`.
8. Markdown output.

Tests use synthetic workbooks generated in test code.

**Verification:** all tests pass. Diffing `Spatac-20261106.xlsx` against
`Spatac-20261121.xlsx` reproduces the known changes:

- the L185 input change and the new comment on it;
- B22 and the A1 date on Medium_Term_Cashflow;
- no formula-text changes;
- recalculated cells reported only with `--values`.

### Phase 3 — Workbook restructure, step A: reproduce

Write the migration script (`tools/migrate/`, Python driving headless LibreOffice). It
produces the restructured workbook from `Spatac-20261121.xlsx`:

- Tabs, as in section 4.1.
- Settings and README tabs.
- `RFS_` tabs with their helper and check columns.
- Expenditures layout, as in section 4.2, including:
  - the RFS and Revised blocks;
  - the Current estimate columns;
  - the migration of column-L formulas;
  - column groups and frozen panes.
- Medium_Term_Cashflow, as in section 4.3, with the planned grid driving the Planned
  section.
- Schedule_of_Investments, as in section 4.4: Held and Planned sections and year columns.
- Projected_Balances, as in section 4.5.
- Defined names (section 3.4), input and calculated styles, and password-less protection
  (section 3.5).
- The clean-up list in section 5.

Today's bugs are preserved in this step.

**Verification:**

- The key results (Medium_Term_Cashflow rows, Projected_Balances balances and minimum
  margins, Expenditures totals) match the original workbook's values after a headless
  recalculation.
- Every formula uses only the functions allowed by section 3.3.
- `xlsx-diff` original → restructured is reviewed. Its report is saved in `_nogit`.

### Phase 4 — Workbook restructure, step B: bug fixes

Apply each bug fix from section 5 as a separate, labelled step: an addition to the
migration script and its own commit. For each one, record the effect on key results
(before and after) in the commit message.

**Verification:** the only result changes are those each fix explains, as confirmed by
`xlsx-diff --values` between consecutive outputs.

### Phase 5 — `spatac-investment-plan`

Implement section 8.3:

- locate content by defined names;
- copy the two tabs;
- turn references to other tabs into cached values;
- remove the sections that are dropped;
- build the "Changes from the RFS" list, with blank-reason flags;
- add the cover block and protection.

Tests run against a synthetic Spatac-shaped workbook generated in test code.

**Verification:**

- All tests pass.
- Running the tool on the phase 4 workbook produces an InvestmentPlan workbook that:
  - opens and recalculates in LibreOffice headless without errors;
  - has totals equal to the source;
  - contains no references to absent tabs.

### Phase 6 — Documentation

- mdBook sources in `docs/`, excluding `docs/archive/`:
  - a chapter per purpose;
  - procedures: rollover, new RFS, parallel edits, saving in either app;
  - a guide for using the CLI tools;
  - an Excel checklist, which also covers the Excel-side checks in section 11.
- README-tab text, added to the workbook by the migration script.
- A GitHub Pages workflow.
- The Investment Plan chapter written so it exports cleanly to PDF.

**Verification:** `mdbook build` succeeds with no broken links. The README tab appears in
the regenerated workbook.

### Phase 7 — Release readiness

- Update `CHANGELOG.md` and set the version.
- Run the full CI command set locally.
- Write a short hand-off note in the final report:
  - what was verified;
  - what still requires the user: the Excel checklist, the Zoom usability judgement,
    pushing, tagging and enabling GitHub Pages.

**Verification:** the CI command set passes locally. The working tree is clean.

## 13. Revisions of 2026-11-22

Decisions from the review session of 2026-11-22, which addressed
`docs/archive/defects-20261121.md`. Question numbers here are prefixed **R** (R1–R26) to keep
them apart from those of the first session; `docs/archive/QA-LOG.md` has the discussion.

Where this section conflicts with sections 1–12, this section wins.

### 13.1 Inputs and sources

- **The review copy** `_nogit/Spatac-20261121-restructured-review.xlsx` is the authority for
  the items the user added or changed in it, and only for those (R2): README and Settings
  text, the Medium_Term_Cashflow row changes (blank row inserted after row 11, closing balance
  moved up to row 20, the rows below renumbered), cell formats on every tab, and on
  Expenditures the title, freeze panes (6 rows, 8 columns), the column groups and hidden
  column F. Everything else stays as the current build makes it.
- All of it is encoded in `migrate.py`, so the workbook is still rebuilt from
  `_nogit/Spatac-20261121.xlsx`. The review copy is a reference only. The user had turned off
  sheet protection to edit it; that is not one of the changes.
- **Outputs** (R16): `_nogit/Spatac-20261122.xlsx` and `_nogit/InvestmentPlan-20261122.xlsx`:
  only the current date in the name (the user found two date stamps too long).

### 13.2 Workbook-wide

- **Vertical alignment: top**, set on the default cell style (Excel's "Normal"), so every tab,
  the RFS tabs included, gets it (R3).
- **Grid borders**: the border the user set on Expenditures J9 goes into the `Input` and
  `Calculated` cell styles, so every input and calculated cell gets it (R13).
- **Protection** (R7): every Spatac tab is protected (no password) when the build ends, as
  before (3.5); the RFS paste-in tabs stay unprotected so a new RFS can be pasted. What
  matters is each cell's own Protected flag: only real input cells are unprotected.
- **Tab order**: Settings moves to the very end, after RFS_cashflow (R6).

### 13.3 README and Settings

- README and Settings take the review copy's text, including the new "Notice for Excel users".
- README's text rows get **explicit row heights** computed from their text, because Excel does
  not re-fit wrapped rows when it opens the file (R5). README stays protected. Rows may look
  loose when zoomed; the user accepts that.
- README describes the **Spatac row** and **RFS check** columns of Expenditures (R17).

### 13.4 Expenditures

- **Protected identity columns** (R7): A–F (row #, item, element, GL account, GL account name,
  Spatac row) are protected and take the non-input look. Adding or removing items means
  unprotecting the sheet; the docs say so. Column A stays visible (in the frozen area) because
  sorting by it restores the original order.
- **2026 Approved** is fixed once the year begins (R9): protected, with a third cell style,
  "fixed for the year". It is filled at rollover with the sheet unprotected (`rollover.md`).
- **Headers** (R8, R17):
  - "Latest estimate" → **"Late-year estimate"** (column M), in the header, the range names
    and the docs.
  - "Actual to date" → **"Actuals to date"**.
  - U6 → **"2027 Reason for change"**. The reasons explain only the next year's approved
    changes. Its position is unchanged and no other reason columns are added.
- **Row 4 totals** are blank for the change-only columns L, M, O, P, R and S (R10). Approved,
  Actuals and the three Current estimate columns keep their totals.
- **Zeros show as 0** in the estimate columns L–T (Mid-year, Late-year and Current estimate
  of all three years), with the number format `#,##0` (until 2026-11-23 `#,##0;-#,##0;0`,
  which showed -0.33 as -0). A blank there means "no change" and a 0 means "will cost nothing", so the two must look different; the old format
  `#,##0;-#,##0;""` showed both as an empty cell (e.g. P29 and Q29). A blank cell still shows
  blank. Approved (J), Actuals (K) and the RFS and Revised blocks keep the zero-hiding format:
  there a zero and a blank mean the same thing. Added 2026-11-22, after phase 14.
- **Current estimates are empty when their inputs are** (N, Q, T): a Current estimate holds
  an empty string, not 0, when every input it chooses from is empty, so that showing zeros does
  not fill these columns with spurious ones. N falls back to Approved (J), which may be blank;
  Q and T fall back to the RFS amount (X, Y), whose lookup gives 0 when the RFS has no amount
  for the item, and that 0 counts as empty. A typed 0 in a Mid-year or Late-year column, or an
  Approved 0, still gives 0. The Revised block passes the empty string on; its totals, the
  variance row, Medium_Term_Cashflow and Projected_Balances read it through `SUM`, which skips
  it. Formerly N used `N($J)`, which turned a blank Approved into 0.
- **Columns A and B left-aligned** (Row # and Item), on every item row, as in the original
  sheet. Numbers otherwise align right, so a right-aligned Row # abutted the Item beside it
  and the two read as a single string. Centring was tried and rejected in favour of the
  original's left alignment.
- **Title in A1**: the review copy's B1 title moves to A1, now that column A is visible.
  Added 2026-11-22.
- **Row 3**: the label leaves C3; "Vendor total" goes in V3, next to the figures (R11).
- **Variance vs RFS** (R12): a header-area row under N, Q, T and each Revised-block year
  (AY–BZ): that column's total minus the RFS total for the same year.
- **Boxes** round O2:Q5 (next year) and R2:T5 (year after), like the user's box round J2:N5
  (R17).
- The **RFS check** column is made narrower (R17).
- **The RFS check works both ways round.** At build every row with a GL account and no RFS
  amounts is marked S, so the check starts clear. A row without S is flagged "Not in RFS" or
  "No RFS amounts" as before; a row with S whose item and element have RFS amounts is flagged
  "In RFS: clear S". Formerly an S exempted the row from every check, so an S made stale by a
  new RFS went unnoticed; its amounts were still used, since the lookup ignores F, but the S
  kept the row from ever being checked again. The S affects nothing but this check.
  `check_blocks.py` toggles an S each way and expects the flags. README's Spatac row text
  now matches the rule (no RFS amounts), not "an item Spatac added". Added 2026-11-22.
- **Comments** (R14): the 12 comments of the old `fvalues_spatac` tab are restored on the
  matching cells of the new layout. Three of them (I3, N221, P221) are Excel threaded comments;
  their text follows "Comment:" after Excel's legacy stub, and only that text is kept (e.g.
  N221: "$42448.32 moved to 2025."). The 17 doubt comments in column U stay.

### 13.5 Block boundaries, sorting and filtering (R1 a, R18–R21)

This replaces the marker-row proposal of the 2026-11-21 handoff.

- **Top boundary: the header row.** The row-block names (`Exp_*`; `Inv_Held`, `Inv_HeldFlows`
  and the per-column `Inv_*` names) start at the block's header row, so a row inserted just
  below the header is inside the name. There is no ▲ start row, because sorting would move it.
- **Bottom boundary: a "▼ End" row** inside the names: grey, labelled "▼ End of items: insert
  new rows above this line" (on Schedule_of_Investments, "… of held investments"). Its row #,
  item, GL account and sort-key cells are blank, so every sort leaves it, and the spare blank
  rows, at the bottom. It holds no amounts.
- The same scheme applies to Expenditures and to the Schedule_of_Investments Held section.
  The Planned section moves down on inserts; its formulas must stay relative (the `AGGREGATE`
  k is `ROWS(M$49:M49)`-style). This is checked after the change.
- Deleting the header row or the ▼ End row breaks the scheme; the docs say so.
- **Filtering**: AutoFilter on the Expenditures header row, allowed while the sheet is protected.
- **Sorting** needs the sheet unprotected, because Excel will not sort a range containing
  protected cells on a protected sheet. The docs give the steps: unprotect, sort, protect.
- **Item sort key**: a helper column in the folded group gives item numbers a sortable form
  (2.1 before 10.1; 18.3.1, 15.5a in place). GL accounts already sort as 4-decimal text.
- **Row #** is typed. An inserted item gets a number between its neighbours, e.g. 125.5, so
  sorting by Row # restores the original order. It is never renumbered automatically.
- Checked in LibreOffice by the build checks; an Excel check is added to
  `docs/procedures/excel-checklist.md`.

**As built (phase 10).** Where these differ from the bullets above, they win.

- Expenditures rows: 3 vendor totals, 4 totals, 5 variance, 6 provenance notes, 7 headers,
  items from 8, then ▼ End (row 250 for the November 2026 workbook). The AutoFilter covers
  row 7 through ▼ End.
- **The years of the RFS and Revised blocks are in row 6**, not the header row:
  `Exp_RFSYears` and `Exp_RevisedYears` point there. The header row is inside the block
  names, and a number there would be added in by every column sum (`SUM(INDEX(Exp_Revised;
  0;k))`). The header cells read "RFS 2026", "Revised 2026" as text.
- On Schedule_of_Investments the top boundary is the "Held" label row (6), which sits between
  the headers and the first held investment; ▼ End is row 47.
- The **old trailing blank item rows are dropped**: with ▼ End, new items are inserted above
  it, and blank rows would sort among the items. The Held section keeps its empty input rows.
- The **item sort key** is column CA (`Exp_SortKey`), in its own folded group, so that no
  existing column moves. Its formula uses `LET` and `FIND`; `FIND` is added to the allowed
  functions (3.3). It pads each dot-separated part to four digits and keeps a letter suffix.
- The AutoFilter protection option cannot be set through LibreOffice's API: the build saves to
  a temporary file and writes the output once with `autoFilter="false"` on Expenditures'
  sheet protection. LibreOffice keeps the option on a round trip. The output must be a new
  file.
- `tools/migrate/check_blocks.py` performs the inserts and sorts on an in-memory copy.

### 13.6 Projected_Balances (R15, R22, R23, R26)

Rebuilt; this replaces section 4.5 where they differ.

- **Two sections over the same year rows**:
  - **RFS baseline**, read-only, from RFS_cashflow.
  - **What-if**, whose contribution increase rates and resulting contribution amounts are both
    visible.
- **Three what-if inputs**, independent of each other:
  - **Start**: "RFS" starts in the RFS's first year from the RFS opening balance; "Current
    year" starts in `Cfg_BaseYear` from a typed opening balance.
  - **Expenditures**: RFS or Revised (from the Expenditures Revised block).
  - **Contributions**: the RFS schedule, or the RFS schedule through a chosen year followed by
    the typed increase rates.
- **Interest** follows the vendor's method: opening balance × the short-term rate for the first
  3 years of the study, then the long-term rate.
- **Validation**: a "difference from RFS balance" column is always visible. With all three
  inputs set to RFS it is zero in every year, which validates RFS_cashflow. The hidden check
  columns (old X–AA) are removed.
- **Opening balance for "Current year"** is a **typed input**, not a formula (R23). The docs
  say what it comprises, all from the financial statements: cash, less payables and accrued
  liabilities, plus investment principal and accrued interest.
- Minimum-balance margins and the "lowest over 30 years" table stay, computed from the what-if
  balance.
- A2 is reworded to describe the tab without the undefined word "Revised" (R17).
- **Past-year actuals are dropped** (R26): the typed actual expenditures and actual closing
  balance per year (old columns F and G) are removed. Comparisons of past years with the RFS
  belong in the prior years' workbooks.

**As built (phase 11).**

- Layout: inputs B4:B12 (dropdowns for B4, B6 and B7), the 30-year lows G4:I10, headers in
  row 14, years 2024–2073 in rows 15–64. RFS baseline B–H, what-if J–S, inflator U (folded).
  Years before the what-if's start year are blank in J–S.
- The current-year opening balance (B5) is **seeded** with the old simulator's projected
  closing balance for 2025, which carried the 2024 audited balance and the 2025 actual
  expenditures; a comment on B5 says to replace it with the figure from the financial
  statements. With it, the default what-if reproduces the old simulator from 2026 on, and
  `verify.py` compares them.
- **Two validations** in `verify.py`, both in memory: all three choices set to RFS; and
  Contributions set to Increases from the study year at the RFS's own increases
  (RFS_CatchUpIncreases), which also checks the vendor's contribution arithmetic (the old
  check column Y). Both must give a zero difference in every year. The docs give both
  procedures.

### 13.7 Investment Plan tool

- **Changes_from_RFS** is removed (R17). The Cover says, in generic terms, that the reasons
  for change from the RFS are described in the Spatac workbooks for the current and previous
  years.
- **Generated date** (R1 b): the cover's date comes from the local clock, not UTC, and a
  `--date YYYY-MM-DD` option overrides it.

### 13.8 Other items (R1)

- `CHANGELOG.md`: 0.1.0 is dated 2026-11-21.
- `xlsx-diff`: a sheet whose protection was turned off is reported as such; its cells are
  not reported as "Cells unlocked" unless their own Protected flag changed.
  *As built (phase 12):* the false reports did not come from the sheet protection.
  umya-spreadsheet drops a cell's protection, including what it inherits from its cell style,
  when the cell's format says `applyProtection="false"`, which is how LibreOffice writes its
  own unlocked input cells. `src/diff/protection.rs` reads the unlocked state from the file's
  formats instead: a format's own `<protection>` decides, else its cell style's, else locked.
- `spatac-investment-plan`: `chrono` (local date) and `xlsx-diff`: `zip` (reading the formats)
  are declared as direct dependencies; both were already linked through umya-spreadsheet, with
  the same features, so no crate was added to the notices.
- `xlsx-diff`: identical changes are grouped across a row as well as down a column.
- **Rollover of Medium_Term_Cashflow (R25)**: the manual one-column-left copy of the grid,
  carry-over and adjustment rows (rollover.md, Medium_Term_Cashflow step 2) is replaced by
  deleting and inserting a column. This replaces the year-heading part of section 4.3.
  - Each year column carries a **fixed, typed year** heading, no longer computed from
    `Cfg_BaseYear`. Every column's formulas are identical and work from their own heading.
  - A grey **"▶ End" column** sits right after the last year, inside the column-block names,
    with the note "insert new year columns to the left of this column". It holds no amounts.
  - A **check cell** warns when the first year heading differs from `Cfg_BaseYear`.
  - At rollover: change the base year; delete the old year's column (typed values move left
    with their columns); insert one column left of "▶ End" and fill it from its left
    neighbour. The window stays six years; the previous dated file is the record of the
    deleted year (4.1).
  - Hiding old columns instead was rejected: it grows the workbook and clutters it.
  - The extractor, Schedule_of_Investments' Planned list and `verify.py` are checked to
    follow the names after a delete-and-insert.
  - *As built (phase 13):* years typed in B3:G3; ▶ End is column H, and every `MTC_*` row
    name runs B–H. The planned-maturities row (11) reads the grid columns to its left from
    column A, which is never deleted (`IFERROR` drops A's text labels), so the same formula
    stands in every year column. The opening cash of the first column stays typed: after the
    delete it shows `#REF!` until the rollover step that types it. The check is A2.
  - Before the delete, the old year's planned investments must all have been bought and moved
    to Held; whatever remains in its grid column is lost. The rollover procedure says so.
  - `check_blocks.py` performs the rollover in memory and checks the years, each year's
    inputs and looked-up figures, the absence of errors, the check cell and the Planned list.
- Excel checklist: add the new items (boundary rows, AutoFilter on a protected sheet, README
  row heights, the what-if validation). The user runs it.
- **GitHub Pages** (the 404): the Documentation workflow builds but cannot deploy because
  Pages is not enabled. The user sets Settings › Pages › Source to "GitHub Actions" and
  re-runs the workflow.

### 13.9 Phases of work

The rules of section 12 apply: work on `aicode`; each phase ends with its verification
passing and a commit; a failed verification or a design change is a stop; nothing
outward-facing without explicit instruction; `_nogit` is create-only.

The common verification, run at the end of every phase that changes the workbook:

- build `_nogit`-bound output in the scratchpad with `migrate.py --fix compounding --fix closing`
  (since 13.10 the fixes are always applied and the options are gone);
- `verify.py --expect-changes` against `Spatac-20261121.xlsx` (since 13.10, plain `verify.py`,
  which lists the fixes' differences as expected). Its expected-difference list is
  updated in each phase for the changes that phase makes on purpose, and each new expected
  difference is named in the commit message;
- `check_investment_plan.py` on a freshly extracted plan;
- the CI command set and `scripts/check-doc-links.py`.

**Phase 8 — Workbook-wide formatting, README and Settings** (13.2, 13.3)
- Default style top alignment; borders in `Input`/`Calculated`; the "fixed for the year" style.
- Settings tab moved last; check that nothing locates tabs by position.
- README and Settings text from the review copy; README row heights.
- The review copy's format changes on Medium_Term_Cashflow, Schedule_of_Investments,
  Projected_Balances and the RFS tabs, and the Medium_Term_Cashflow row changes (rows
  renumbered; `verify.py`, the extractor and the docs follow the names, not row numbers).
- Protection of every sheet at the end of the build.

**Phase 9 — Expenditures** (13.4)
- Review copy: title, freeze panes, groups, hidden F, formats.
- Protected A–F and Approved; header renames, including range names; row 4 totals; row 3
  label; Variance vs RFS row; year-block boxes; RFS check width; restored comments.

**Phase 10 — Block boundaries, sorting and filtering** (13.5)
- Names start at the header rows; ▼ End rows on Expenditures and the Held section.
- Item sort key column; AutoFilter allowed under protection.
- Check: insert a row just below each header and just above each ▼ End row, then confirm
  every name grew and totals, lookups and the Planned list are unchanged. Sort by GL and by the
  item key, then by Row #, and confirm the workbook returns to its original values.

**Phase 11 — Projected_Balances** (13.6)
- Rebuild the tab; the extractor and `check_investment_plan.py` follow if they read it.
- Check: with all three inputs set to RFS, the difference column is zero in every year.
  `verify.py`'s Projected_Balances key results are replaced by this check and by the
  current-year projection.

**Phase 12 — Tools** (13.7, 13.8)
- `spatac-investment-plan`: Changes_from_RFS removed, Cover text, local date and `--date`;
  `tests/investment_plan.rs` updated.
- `xlsx-diff`: the protection report fix and row-wise grouping, with tests in `tests/diff.rs`.
- `CHANGELOG.md`: 0.1.0 date; an Unreleased section listing this session's changes.

**Phase 13 — Rollover grid** (13.8)
- Fixed year headings, the "▶ End" column and the base-year check cell (13.8, R25).
- Check: on a scratch copy, perform the rollover (base year +1, delete column B, insert and
  fill a column before "▶ End"), then confirm the remaining years' figures, the Planned list
  and the extracted plan are unchanged apart from the year shift, and the check cell is clear.

**Phase 14 — Documentation and hand-off**
- Update `docs/reference/names.md`, `docs/purposes/track-projects.md` ("Adding a row",
  sorting and filtering), `docs/purposes/investment-plan.md`, the Projected_Balances purpose
  pages, `docs/procedures/rollover.md` (Approved, the Held section, the grid) and
  `docs/procedures/excel-checklist.md`.
- Write the final outputs to the two `_nogit` names of 13.1.
- A short hand-off note: what was verified, and what still requires the user (the Excel
  checklist, Pages, pushing).

### 13.10 Simplification (2026-11-22)

After phase 14 the user found the workbook brittle and some of its logic hard to understand.
Agreed changes:

1. **Item sort key removed** (Expenditures column CA, `Exp_SortKey`): a 1,173-character
   formula. Sorting by Row # restores the original order, which is item order.
2. **RFS check** (Expenditures I) rewritten with `LET` and `IFS`, over several lines.
3. **Planned maturities** (Medium_Term_Cashflow row 11) rewritten with `LET` naming the
   amounts, purchase years and terms. Summing the Planned list instead, as first proposed,
   was rejected: the list reads the 5-year balancing row, which depends on row 11, so it would
   be a circular reference.
4. **Planned list** (Schedule_of_Investments rows 49–78): one fixed row per grid cell, by a
   typed Grid # in column M (1–5: the five terms in the first year column, 6–10 the second,
   and so on), blank when the cell is empty. It replaces the `AGGREGATE` formula that packed
   the non-empty cells to the top. At rollover each row shows the cell now in its place. After
   the old year's column is deleted, its opening cash shows `#REF!` until it is typed, and the
   error carries through the 5-year balancing row and the closing balances into the Planned
   5-year rows; the rollover procedure says so. (The last year column's rows read the ▶ End
   column meanwhile, which is empty, so they are blank, not errors.)
5. **Projected_Balances what-if** rebuilt on the old simulator's pattern, with helper columns:
   J Opening balance, K Expenditures, L Contribution increase, M Contributions, N Interest
   rate, O Interest, P Closing balance, Q Difference from RFS balance, R–T margins, U per unit
   per month; W inflator (folded). Before the start year the Opening balance is blank and the
   rest of the row with it, so no formula repeats a "started?" test. The starting balance is
   worked out once, in D5 (`PB_StartBalance`). Years before the start were not filled with the
   RFS figures: the user found that confusing.
6. **Nested IFs** replaced by `IFS` or `SWITCH`: the Current estimates N, Q, T; the Revised
   block (`SWITCH` on the year's distance from the base year); the investments' yearly cash
   flows; the Medium_Term_Cashflow year check (A2).
7. **The phase 4 fixes are permanent.** `migrate.py` has no `--fix` options and no code that
   rebuilds the old bugs; `verify.py` lists the 13 differences the fixes cause, each by its
   cells, and has no `--expect-changes`. On 2026-11-22 four builds were made without the
   options and so carried the old bugs; this makes that impossible.

The results are unchanged: `verify.py` finds only the expected differences, and the what-if
still reproduces the RFS with its choices set to RFS.

After this, one final build is to become the master workbook, changed by hand in Excel from
then on, and `tools/migrate/` is to be retired.

## 14. Revisions of 2026-11-24

Decisions from the session of 2026-11-24, on `docs/archive/Prompt-20261124.md`. Question
numbers are prefixed **G** (G1–G35). Where this section conflicts with sections 1–13, this
section wins.

### 14.1 Sources and method (G5, G35)

- **The golden source** is `_nogit/Spatac-L-20261123.xlsx`, the user's own workbook in the old
  layout. `_nogit/Spatac-20261123c.xlsx` is the current build aligned with it, and is the
  starting point of this restructure. All other files are history.
- **Numeric parity with the golden source is required**, except for the one inflation fix in
  14.4 (G24, G32).
- The restructure is a new scripted step that takes `Spatac-20261123c.xlsx` as input, with
  checks. It does not rebuild from `Spatac-20261121.xlsx`.
- After the user accepts the result it goes to the Board. Their feedback is applied by
  whatever method suits each change, scripted or by hand (G35). The freeze in 13.10 waits
  until then.

### 14.2 Investment Plan tool (G2)

The reference is `_nogit/Medium_Term_Reserve_Cashflow-20261123.xlsx`, a copy the user
formatted by hand. The plan is a frozen document, so:

- Medium_Term_Cashflow and Schedule_of_Investments lose the input fill (tan) and the dotted
  input borders.
- Every tab uses the source's font (Arial 10). The Calibri 11 of today's output is a side
  effect of re-serialising the styles. The year headings are bold.
- Column widths stay as the tool makes them. The reference copy's wider columns were not
  intended.
- The Cover keeps the input look on the cells to be edited before publishing.
- Only the first tab (Cover) is selected when the file opens. Today the Cover is active and
  Medium_Term_Cashflow is also selected.

### 14.3 Medium_Term_Cashflow (G3)

- A blank row is inserted after "Adjustment to minimum cash balance" (row 23). Everything below
  moves down one row, with the formulas and the `MTC_*` names.
- The minimum row (now 25) is relabelled **"Minimum closing cash balance"**, dropping
  "(Settings, inflated)", and is no longer bold.

### 14.4 Adjustments tab (new; G6–G11, G16, G18, G23, G25, G26, G32, G34)

Placed after Expenditures. It records every change to the RFS, with its reason. It replaces
the next-year and year-after estimate blocks and the Reason for change column of Expenditures.

**Blocks.**

- Four permanent blocks, titled "Approved in *YYYY*" with *YYYY* = `RFS_StudyYear` + 0, +1, +2
  and +3 (2024–2027 today). The blocks always exist.
- At the January rollover the user is directed to the next block. At a new RFS all four
  blocks are emptied, and the user starts again in the first block (year 0 of the new study).
- All blocks apply together, and their adjustments add up.
- Each block has a header row, **10 input rows filled in advance** (formulas included), and a
  closing row. More rows may be inserted. The docs say how to copy the formula columns into an
  inserted row, and `spatac-check` reports any missing formula (14.8). No Excel Table is used,
  because the workbook must work in Calc as well as Excel.

**Columns.**

| Col | Content | Kind |
|---|---|---|
| A | Item | Input; dropdown |
| B | Year from | Input; dropdown: "none" or a year |
| C | Year to | Input; dropdown: "none" or a year |
| D | Amount | Input |
| E | Reason for change | Input |
| F | From-year amount | Formula |
| G | To-year amount | Formula |
| H | Check | Formula |

**Meaning of a row.**

- **Move**: both years given. D is in **from-year dollars**. F = D. G = D × factor(C) ÷
  factor(B) (14.5).
- **Add**: Year from is "none". D is in **to-year dollars**. F is empty and G = D.
- **Remove**: Year to is "none". F = D and G is empty.
- A move may go backwards in time; the factor then deflates.
- These rules go into the user documentation (G10).

**Validation.**

- Item: a list from `Expenditures!J` (the adjustment label, 14.6). Section-heading rows give
  blank entries in the list; H rejects any Item that is not a label (G26).
- Year from and Year to: "none" or a year from `RFS_StudyYear` to `RFS_StudyYear` + 29
  (2024–2053).
- Amount: `ISNUMBER` and greater than 0, so a typed number or a formula is accepted.
- Item and Reason are required.

**Check (H).** Empty for an empty row. Otherwise it flags:

- a missing Item, Year from, Year to, Amount or Reason;
- an Item that is not a label;
- Year from equal to Year to;
- an overdraw: F greater than the item's RFS amount in Year from, as adjusted by the rows above
  it (all blocks, in order).

**Initial data.** The user's six rows in the 2026 block of
`_nogit/Spatac-20261123c-restructured.xlsx` (A21:E26) are real data, keeping the user's reasons.
Their amounts are restated by the rules above:

| Item | From → To | Amount |
|---|---|---|
| 3.9 Replace Garage Ramp Topping | 2025 → 2028 | 229,500 |
| 6.3 Replace Exterior Sealants | none → 2027 | 50,000 |
| 6.3 Replace Exterior Sealants | 2026 → 2028 | 416,160 |
| 18.3.1 Refurbish Corridors | 2028 → 2027 | 524,979.60 |
| 18.3.1 Refurbish Corridors | none → 2027 | 85,314.12 (600,000 less row 4 in 2027 dollars) |
| 18.11 Refurbish Boardroom | 2027 → 2025 | 42,448.32 |

- **3.9 (G32):** the golden source had 238,772 in 2028, which carries two years of inflation
  instead of three. The fix gives 229,500 × 1.02³ = 243,547.24. A comment on 3.9's 2028 cell in
  Expenditures records the difference. This is the only departure from numeric parity.
- **18.11 (G25):** kept as the user entered it. The money was probably spent in 2025, and its
  appearance in the 2026 Mid-year estimate was probably an error. The row changes no total.
- **5.1, 5.5, 5.6 and 18.4 (G24):** their 2026 RFS amounts are absent from the current year,
  as in the golden source. No adjustments are added.

### 14.5 Inflation factor (G9)

- The RFS has a **short-term** rate (`RFS_ShortInflation`) for its first years and a
  **long-term** rate (`RFS_LongInflation`) after them. Both are 2% today, which is a
  coincidence; nothing may assume they are equal.
- The number of short-term years appears in the RFS only as a text label ("Inflation 0-3"), so
  it is a new Settings parameter, **Short-term inflation years** = 3 (`Cfg_ShortInflationYears`).
- factor(year) = (1 + short)^MIN(n, k) × (1 + long)^MAX(n − k, 0), with n = year −
  `RFS_StudyYear` and k = `Cfg_ShortInflationYears`.
- The Adjustments tab and the Projected_Balances inflator (`PB_Inflator`, today short-rate
  only) both use this factor.

### 14.6 Expenditures (G12–G15, G17, G22, G23, G27)

**Columns.**

| Cols | Content |
|---|---|
| A–I | Unchanged: identity, keys, RFS check |
| J | **Adjustment label** (formula, protected) |
| K–O | Current year, today's J–N: Approved, Actuals to date, Mid-year, Late-year, Current estimate (`Exp_Current` = O) |
| P | Working notes (today's V) |
| Q | Blank, narrow |
| R | **Next year**: that year's column of `Exp_Adjusted` |
| S | Blank, narrow |
| T–AW | **Adjusted RFS**, 2024–2053, one column per year (`Exp_Adjusted`, `Exp_AdjustedYears`) |

Removed: today's O–Q (next year), R–T (year after), U (Reason for change), W–AX (RFS block) and
AY–BZ (Revised block), with their names (`Exp_MidEst1`…`Exp_Current2`, `Exp_Reason`, `Exp_RFS`,
`Exp_RFSYears`, `Exp_Revised`, `Exp_RevisedYears`).

**Adjustment label (J).**

- Core: GL account + " - " + Building element (C) when the element is not empty; otherwise
  GL account + " - " + GL account name (E), even if the name is empty.
- Spatac rows (F = "S") add " (Spatac)".
- Spaces are trimmed. Rows with no GL account (section headings) give an empty label.
- Data fixes: row 34 (4.1a) gets its GL account as a number (9250.1010, not the text
  '9250.101); row 241 (18.24a) loses the line break in its GL account name.
- Item 16.1 (Garbage Compactor) keeps the RFS number 16.1, not the golden source's "16".

**Adjusted RFS (T–AW).**

- Each cell = the item's RFS amount for the year (from `RFS_fvalues`, by `Exp_Key`)
  + `SUMIFS` of Adjustments G for the item's label with Year to = the year
  − `SUMIFS` of Adjustments F for the item's label with Year from = the year.
- A cell is highlighted by conditional formatting when it differs from the RFS amount by more
  than 0.5.
- Years before `Cfg_BaseYear` are grouped and folded. At rollover the year just ended is
  folded too (G27).

**Next year (R)** reads `Exp_Adjusted` at `Cfg_BaseYear` + 1. There are no next-year estimates:
every next-year change goes through Adjustments (G17).

### 14.7 What reads Expenditures (G14, G22)

- `Medium_Term_Cashflow!B9:G9` (Revised expenditures): for the base year, `-SUM(Exp_Current)`;
  for every other year, minus the year's column sum of `Exp_Adjusted`.
- The Medium_Term_Cashflow RFS row (29 after 14.3) reads the year's column sum of
  `RFS_Fv_Values` (`RFS_Fv_Years`). The totals equal those of the removed RFS block.
- `Projected_Balances!K15:K64` (what-if expenditures): the base year from `Exp_Current`, which
  overrides the adjusted figure; other years 2024–2053 from `Exp_Adjusted`; 2054–2073 from the
  RFS cost in `Projected_Balances!B`. Rows before the what-if start year stay blank.

### 14.8 Notes and comments (G28–G31)

- **Working notes** of the golden source are kept, on their rows, in `Expenditures!P`. Column
  letters in them are rewritten for the new layout (golden I, J, K, L → K, L, M, N), and the
  user approves every changed note before the build. Letters that were already wrong in the
  golden source are translated, not corrected.
- **Cell comments** stay comments, on the cell they belong to:
  - kept: "Kerry's projection" on the Approved header; "Revised by Bruce…" on 15.5a's
    Late-year estimate;
  - dropped: the comments on the removed next-year and year-after columns. Those on the
    Mid-year and Late-year cells are superseded by the Adjustments reasons;
  - added: the 3.9 inflation note (14.4).
- **Claude's texts are removed:** the 55 reasons of `tools/migrate/reasons.py` (column U) and
  the 17 doubt comments on them. The reasons for change are the user's, in Adjustments.
- The golden source's comments on `cashflow_spatac_simulator!H15` and
  `fvalues_spatac_end_2025!H3` are obsolete and ignored.
- `Medium_Term_Cashflow!J33`: the updated reference to Projected_Balances stays.

### 14.9 Where documentation lives (G33)

One place for each kind of information:

- **Docs site** (`docs/`): how the workbook works and how to operate it. Procedures (rollover,
  new RFS, adding items, copying the Adjustments formula columns), names and formulas.
- **README tab**: the purpose of each tab, the input legend, and a link to the docs site. No
  procedures.
- **Working notes** (`Expenditures!P`): the user's notes on an item's figures.
- **Adjustments reasons** (`Adjustments!E`): why the RFS was changed.
- **Cell comments**: only the provenance of a single figure. Never instructions. Claude's four
  instruction comments (`Medium_Term_Cashflow!H3`, `Projected_Balances!B5`, `RFS_info!K1`,
  `RFS_fvalues!BC1`) move into the docs.
- **Check cells** say what is wrong and name the docs page that explains it.

### 14.10 `spatac-check` (G20)

A third binary in the crate. It reads formulas, names, formats and stored results, without
recalculating, and reports every problem it finds, by tab and cell, before exiting. It exits
non-zero if it found any. It checks:

- missing formulas in formula columns (the Adjustments F–H rows, Expenditures);
- defined names that are `#REF!` or out of place;
- stored error results (`#N/A`, `#REF!`, `#DIV/0!`…) from the last save in Excel or Calc;
- missing data validations on the Adjustments input rows;
- a cell's Protected flag that does not match its role.

### 14.11 Order of work

1. Commit `docs/archive/Prompt-20261124.md` and this section.
2. Investment Plan formatting (14.2) and the Medium_Term_Cashflow row (14.3).
3. The restructure (14.4–14.8), with checks.
4. `spatac-check` (14.10).
5. Documentation (14.9), including the rollover and new-RFS procedures for the Adjustments
   blocks.

The rules of section 12 apply: each step ends with its verification and a commit; a failed
verification or a design change is a stop; nothing outward-facing without explicit
instruction; `_nogit` is create-only.

### 14.12 As built

Where these differ from 14.1–14.11, they win.

- **Script.** `tools/migrate/restructure.py` edits `Spatac-20261123c.xlsx` in place and reads the
  user's six rows from the mock-up; `verify_restructure.py` checks it (14.1). Both open copies of
  their sources, so LibreOffice leaves no lock file in `_nogit`.
- **Titles.** Expenditures A1 is "ADJUSTED EXPENDITURES" and Adjustments A1 "ADJUSTMENTS TO
  PLANNED EXPENDITURES", as in the mock-up.
- **Adjustments layout.** Blocks start at rows 4, 20, 36 and 52: title, blank, header, 10 input
  rows, grey closing row, two spacers. The `Adj_*` names run from the first header row (6) to
  the last closing row (65). The Year from and Year to choices are in hidden column J
  (`Adj_YearChoices`). A2 counts the flagged rows. Inserting a row needs the tab unprotected.
- **Inflation.** The factor of 14.5 is computed once, in `PB_Inflator`; Adjustments G divides
  the to-year factor by the from-year factor there. With both RFS rates at 2%, the inflator's
  values, and so every Projected_Balances margin, are unchanged.
- **Expenditures rows 3–5** keep their roles over the new columns: row 3 the vendor's total per
  adjusted year, row 4 the totals (Approved, Actuals, Current estimate, next year, adjusted
  block), row 5 the variance against row 3. Over the adjusted block, row 6 holds the years.
- **The Medium_Term_Cashflow RFS row** (29) reads `RFS_Fv_Totals`, the vendor's own yearly
  total: `RFS_Fv_Values` spans the tab's header rows too, so a column sum would add the years.
- **Projected_Balances K** uses the RFS cost for years before the base year, as well as after
  the 30-year span: past years do not affect the projection (G27).
- **Labels** use the GL account as typed (text, e.g. 9210.2030). 4.1a's account becomes the text
  9250.1010.
- **Spacer and groups (2026-11-25).** At the user's request, the label column J is folded on its
  own, and a narrow blank column K separates it from the current year: in `.xlsx`, groups that
  touch merge into one. The current-year inputs are no longer grouped. Final layout: J label,
  K spacer, L–P current year, Q working notes, R spacer, S next year, T spacer, U–AX adjusted
  RFS. This supersedes the letters of 14.6.
- **Working notes name columns by heading, in quotes** (2026-11-25, superseding the letter
  translation of 14.8): golden I, J, K, L become "Approved", "Actuals to date", "Mid-year
  estimate", "Late-year estimate", so that a note stays right when columns move. The two
  notes that were already wrong (11.5, 18.11: "K = 0: not in YE projection") now say
  "Late-year estimate" = 0.
- **Defects of 2026-11-24** (`docs/archive/defects-20261124.md`):
  - "Vendor total" is now **RFS total**, with legends "RFS total", "Adjusted total" and "Variance vs RFS"
    right-justified in Q3:Q5. Next year (S3) shows its RFS total too.
  - **Variances are RFS minus Spatac** on Expenditures (row 5), as on Medium_Term_Cashflow: a
    negative variance is unfavourable and shows in red.
  - **Years of the adjusted block** stay numbers in row 6 (`Exp_AdjustedYears`), green
    (Calculated) and invisible (format `;;;`); the headings in row 7 are their text (`=""&U$6`),
    which the column totals do not add. At the user's request, a comment says "Do not modify
    the cells to the right of this one in this row…": an exception to the rule of 14.9 that
    comments hold no instructions. The comment (T6) and the block's bold description (T2,
    running across the empty cells to its right) are in the spacer T, which is never folded:
    the block's past years are, and the first of them changes at each rollover. S2 shows only
    the next year, bold; S4 and U4:AX4 are bold; S is wider.
  - **Blank versus zero** (as 13.4): next year and the adjusted block are blank where the RFS has
    no amount and no adjustment touches the year, and show `0` where adjustments leave nothing.
  - **Adjustments Check** shows OK for a non-empty row in order, right-justified; A2 counts the
    other non-empty results.

## 15. Revisions of 2026-11-25

Decisions from the session of 2026-11-25, on `docs/archive/defects-20261125.md`. Question
numbers are prefixed **H** (H1–H19). Where this section conflicts with sections 1–14, this
section wins. The changes go into `tools/migrate/restructure.py` (H2), which still starts from
`_nogit/Spatac-20261123c.xlsx`; `verify_restructure.py` and the docs follow.

### 15.1 Settings (H16, H17, H19)

- A new row 9, **"Years of short-term interest in the RFS (the long-term rate applies after
  them)"**, B9 = 3, an input named `Cfg_ShortInterestYears`. The note moves to row 10.
- Row 8 (`Cfg_ShortInflationYears`) is unchanged: the two counts are separate parameters.
- Interest is at `RFS_ShortInterest` for the years `RFS_StudyYear` to `RFS_StudyYear` +
  `Cfg_ShortInterestYears` − 1 (2024–2026), then `RFS_LongInterest`. Formerly the what-if used
  the short rate through `RFS_StudyYear` + 3 (four years). Both rates are 4% today, so no
  number changes.

### 15.2 Projected_Balances: the RFS block is recomputed (H8–H11)

This replaces the "RFS baseline, read-only" of 13.6.

- **Read from the RFS:** only the expenditures (RFS_cashflow B), the rates (interest G7:G8,
  contribution increases K9:K12) and two starting values: the opening balance (C11) and the
  present contribution (C12). **Everything else is recomputed**: contributions (last year's
  × (1 + the RFS increase)), interest (opening × rate), closing balance.
- **Same columns as the what-if block**, so the two read across column for column. The column
  where the what-if has "Difference from RFS balance" holds the **check**: the recomputed
  closing balance minus RFS_cashflow G. A difference over 0.5 is shaded; the count of such
  years is in row 13 above the column (red when not zero).
- Checked outside the workbook before building: the recomputation matches RFS_cashflow C, D
  and G within 0.01 in all 50 years.

### 15.3 Projected_Balances: the what-if is simpler (H13–H15)

- The three "RFS" choices go: **start** (always the current year, `Cfg_BaseYear`, from the
  typed opening balance), **expenditures** (always Revised: `Exp_Current` for the base year,
  `Exp_Adjusted` after, the RFS beyond 2053) and **contributions** (always from increases).
  `PB_Start`, `PB_ExpBasis`, `PB_ContribBasis` and `PB_StartBalance` are removed. The check of
  the vendor's arithmetic is now the RFS block's (15.2); A3's instruction goes.
- **"Contribution increases follow the RFS schedule through"** (formerly "Contributions follow
  …"): through that year the increase column shows the RFS's increase, after it the three
  catch-up increases and then the subsequent one. Contributions are always last year's × (1 +
  increase), the first what-if year growing from the RFS block's recomputed contribution of
  the year before. With the year set to 2073 the what-if follows the RFS schedule throughout.
- This removes the blanks in the increase column for the years through that year (L17:L18 in
  `Spatac-20261124f.xlsx`), which were by design, not a formula error (H14).
- "Subsequent annual increase" stays an input whose default is `=RFS_SubsequentIncrease` (H18).
- "Difference from RFS balance" compares with the recomputed RFS balance.

### 15.4 Projected_Balances: layout and format (H3–H7)

- **Inputs** A4:B10 (moved to P4:P10 by 15.5): what-if start year (calculated), opening
  balance, the "through" year, the three catch-up increases, the subsequent increase. The old D8:E9 (a historical inflation
  rate and an unlabelled value of the old simulator; no formula used either) are deleted, here
  and in `migrate.py`.
- **Lowest margins over the first 30 years**: label in D, amount in G, year in H, rows 4–10
  (replaced by 15.5).
- **Year table**, header row 14, years in rows 15–64: A year; B–M the RFS block; N spacer;
  O–Z the what-if; AB the inflator (folded).
- **Labels**: "PV margin over minimum" replaces "margin over minimum (PV)" everywhere on the
  tab ("RFS PV margin over minimum", "What-if: PV margin over minimum", …).
- **Negatives in red** with a minus sign (`MONEY_RED`) in every money cell of the tab.
- **Years past the 30-year horizon** (rows 45–64) have a light grey fill across the table.
- Column widths: every table column fits a year; there is no separate year spacer to widen
  (H4).

### 15.5 Review of `Spatac-20261125a.xlsx`

- **Inputs above the what-if**: A4:B10 move to O4:P10. The labels are right-justified in O and
  run left over the empty cells before it. Column A holds only years now and is narrowed.
- **Two summaries**, one above each block, over the first 30 years: B4:G7 for the RFS
  (`PB_RFSLowest`, B5:G7) and R4:W7 for the what-if (`PB_Lowest`, R5:W7), the figures starting
  two columns after the label (D and T). Each has a row per
  margin (over minimum, over inflated minimum, PV) and the columns Lowest, Year, **First
  negative**, Year. The first negative year is `MINIFS` of the years where the margin is below
  zero; "none" when there is none, with the amount blank.
- **`-0`**: superseded by 15.6.

### 15.6 Money format

- **Projected_Balances only** uses the legacy workbook's red format, `#,##0;[RED]-#,##0`
  (`MONEY_RED`; with cents, `#,##0.00;[RED]-#,##0.00` for the per-unit columns), in every money
  cell. The other tabs keep their formats as they were (`#,##0`, the hide-zero `#,##0;-#,##0;""`
  of 13.4, and red only on their variance rows). Making every tab red was tried and rejected
  by the user as too much red noise.
- **Expenditures Approved and Actuals to date** (L8:M249) change from the hide-zero format to
  `#,##0` (`MONEY_ZERO`): a zero shows 0, black. This supersedes 13.4's exception for them.
- **Expenditures totals** (rows 3–4 of next year and the adjusted block) also change from the
  hide-zero format, carried over from the old RFS and Revised blocks, to `#,##0`: a year's total
  of 0 is a result, not an absence. Every amount on Expenditures is now `#,##0`, except the red
  variance row 5.
- **No hidden zeros anywhere.** The hide-zero format of 13.4 is gone from every Spatac tab: the
  last cells with it, the Schedule_of_Investments cash-flow grid (N7:X80), take `#,##0`, so its
  empty years show 0. `verify_restructure.py` checks that no cell hides a zero.
- **`-0` is accepted.** An amount between −0.5 and 0 shows `-0`, as in the legacy workbook. A
  conditional format (`[RED][<=-0.5]#,##0;#,##0`) would avoid it, but the user prefers the
  simple one. Today only Projected_Balances I26, I27 and I29 (about −1e-8, rounding noise in the
  check) show it.

### 15.7 Projected_Balances rows

- The blank row 11 is deleted: the block headings and the "years that differ" count are in row
  12, the column headings in row 13, the years in rows 14–63 (the light grey years 44–63), and
  the panes freeze below row 13. This supersedes the row numbers of 15.2–15.6 (the count in
  I12, the `-0`s in I25, I26 and I28).

## 16. Revisions of 2026-11-26

Decisions from the session of 2026-11-26, on `docs/archive/defects-20261126.md` and
`docs/archive/Prompt-20261126.md`. Question numbers are prefixed **I** (I1–I21). Where this
section conflicts with sections 1–15, this section wins.

### 16.1 Baseline and build (I1–I3)

- The baseline is the user's `_nogit/Spatac-20261125h-PV.xlsx`: `Spatac-20261125h.xlsx` with
  the Medium_Term_Cashflow fix from the DS review (the 2027 one-year investment 320,000; the
  2028 adjustment to minimum cash balance 311), which removes the negative 2028 five-year
  investment.
- A new script, `tools/migrate/update_20261126.py`, edits a copy of the baseline in place, as
  `restructure.py` did. `restructure.py` is no longer re-run: it would have to replay the
  user's hand edits. `verify_20261126.py` requires every value to equal the baseline's except
  where this section changes it on purpose.
- From the user's `_nogit/Spatac-20261125h-PV-edits.xlsx` only display changes are taken: the
  Projected_Balances column headings B13:M13, the Settings A6 label and a blank row before the
  Settings note. Its what-if increases (P7 4%, P8:P9 `=P7`) and its unprotected tabs are not.
- `design.md` is now `design-history.md`; the scripts' references follow.

### 16.2 README and row heights (I4–I6)

- Calc and Excel fit a row's height at 100% zoom; at another zoom the text can reflow onto one
  more line and be cut off (README A3 at 110%). Every README line has an explicit height: the
  fitted height with Excel's 10% margin, plus one line of slack when the text ends within 15%
  of a line's end. Each Adjustments block header gets the fitted height plus one line.
- The documentation URL moves to a row of its own (A4), below its description.
- Each line's topic, the text before " — ", is bold. The text edits of the defects file are
  made; the RFS check's "Check it after pasting a new RFS or adding a row." becomes "The column
  updates by itself. After pasting a new RFS or adding a row, scan it and resolve every flag."
- The colour legend shows each colour: Orange and Light grey take their fills (they had the
  Grey style), and a new line describes the Red alarm (16.3).

### 16.3 The Alarm format (I8, I12)

- A failed check shows a **red fill with bold white text** (cell style `Alarm`, fill
  `0xC00000`), applied by a conditional format: a count that is not 0, and each value it
  counts. It replaces the orange shading of Projected_Balances column I and the red font of
  I12.
- Medium_Term_Cashflow's investment grid (B14:G18) uses it for a negative investment.

### 16.4 The RFS totals are recomputed and checked (I7, I16–I21)

- New name `RFS_Fv_Items` = RFS_fvalues C15:AZ400, the item rows below the "Item | Building
  Element" header. Section headings in it (item 0) hold 0. In the baseline the items add up to
  the vendor's "Total Annual Expected Costs" (row 10) in all 50 years.
- **Expenditures row 3** (U3:AX3, "RFS total (sum of items)") adds up the items instead of
  reading row 10: `=SUM(INDEX(RFS_Fv_Items,0,MATCH(U$6,RFS_Fv_Years,0)))`. Each year is
  flagged (16.3) where it differs from row 10 by more than 0.5; P3 counts those years
  (`SUMPRODUCT` over `SUMIF`), labelled in O3 "Years where items differ from RFS_fvalues
  totals". No row is added.
- **Projected_Balances column C** (Expected costs) adds up the items too, for all 50 years:
  `=-SUM(INDEX(RFS_Fv_Items,0,MATCH($A14,RFS_Fv_Years,0)))`. Each year is flagged where it
  differs from RFS_cashflow's expected cost; F12 counts them, labelled "Cost years that
  differ" in E12. No column is added, and no formula on Projected_Balances reads
  Expenditures. (An earlier proposal, a count across all items and years in one cell, needed
  `MMULT` and an array formula, which Excel 2019 would enter with Ctrl+Shift+Enter; computing
  C from the items removed the need.)
- To fit both counts in row 12: the block heading B12 becomes "RFS (recomputed)", and H12
  "Balance years that differ".

### 16.5 No hidden errors; the what-if is independent (I10, I11, I19)

- Rule: a formula never turns an error into a value. Removed:
  - Projected_Balances C and I: `IFERROR` around lookups into RFS_cashflow, which covers all
    50 years.
  - Projected_Balances P: `ISNA(adjusted_column)`, which also hid a year missing inside the
    adjusted block. It is now an explicit `year>MAX(Exp_AdjustedYears)` test; those years
    (2054–2073) use the sum of the RFS items.
  - Expenditures U3:AX3: `IFERROR(…,0)` (replaced by 16.4).
  - Adjustments G (To-year amount): the outer `IFERROR(…,"")`. Its blank cases are handled
    inside; it could only hide a year missing from `PB_Year`, which the validation prevents.
- Kept, each because the "not found" case is a legitimate value, not an error:
  - Expenditures adjusted amounts and their orange conditional format: an item with no RFS
    amounts (a Spatac row) is 0.
  - Expenditures H (Key2): a blank or non-numeric GL account gives an empty key; the RFS check
    then flags the row.
  - Adjustments H (Check): an item not in Expenditures becomes the flag "Item not in
    Expenditures"; an item with no RFS amounts holds 0.
  - Medium_Term_Cashflow row 11 (Planned investments maturing): `IFERROR` inside `SUMPRODUCT`
    drops column A's text labels (13.8).
  - RFS_info J and K: a non-numeric GL cell is a heading of the vendor's sheet; a key not on
    Expenditures becomes the flag "No Expenditures row".
- **The what-if reads no cell of the RFS block** (B–M). P reads `RFS_Fv_Items` beyond the
  adjusted block (was `$C`); Q reads `RFS_CatchUpIncreases` and `RFS_SubsequentIncrease`
  through P6 (was `$D`); R's first year grows from `RFS_Cf_Contrib` of the year before (was
  `$E`); S reads `RFS_ShortInterest` and `RFS_LongInterest` (was `$F`). Only V, a comparison,
  reads H. Q, R and S are `LET` and `IFS`, with no nested `IF`. Every what-if value is
  unchanged; the verifier clears B:M and checks that O:Z (bar V) stay the same.

### 16.6 The what-if summary span (I9, I18)

- New Settings input B10, `Cfg_YearsBetweenRFS` = 3, "Years between RFS updates"; a blank
  row 11; the note moves to A12.
- The what-if summary (R4:W7) runs from the current year through the RFS's last year plus
  `Cfg_YearsBetweenRFS`: 2026–2056 today, the span of the next RFS. R4 shows it: "What-if,
  2026–2056". Lowest, Year and First negative use `MINIFS` and `COUNTIFS` over the years in
  the span. The RFS summary (B4:G7) keeps the first 30 years.
- Today: lowest margin over minimum −2,115,795 (2055), first negative −1,058,672 (2045).

### 16.7 Medium_Term_Cashflow (I12, I15)

- The investment grid (B14:G18) flags a negative investment (16.3).
- A comment on A23 explains the plug "Adjustment to minimum cash balance"; the Investment Plan
  page explains it in full, by headings rather than row numbers.

### 16.8 Projected_Balances layout (I13, I14)

- The RFS block (B14:M63) has no fill, which sets it apart from the green what-if; 2054–2073
  stay light grey.
- The Inflator column (AB) is no longer grouped and always shows.
- Row 14 (2024) had kept the header's height since 15.7; it is fitted.

## 17. Revisions of 2026-11-27

Decisions from the session of 2026-11-27, on `docs/archive/defects-20261127.md` and
`docs/archive/Prompt-20261127.md`. Question numbers are prefixed **J** (J1–J12). Where this
section conflicts with sections 1–16, this section wins.

### 17.1 Baseline and build (J1, J7)

- The baseline is the user's `_nogit/Spatac-20261126a-edits.xlsx`: the previous build with their
  edits. On Expenditures they inserted a column at T: the RFS totals count moved from P3 to T1,
  labelled in U1 "  No. of RFS years where totals don’t add-up correctly" (leading spaces kept,
  J5), O3:P3 emptied, and the RFS block moved from U:AX to V:AY. On README they shortened the
  Red legend line (A21) and bolded A25 and A28; on Adjustments they reworded the amount headings
  (D6, D22, D38, D54) to "Amount in from-year $s (in to-year $s when Year from is none)"; on
  Settings they widened column A and cleared B11's input format. Their unprotected tabs are not
  kept.
- `tools/migrate/update_20261127.py` edits a copy of the baseline in place, as
  `update_20261126.py` did. LibreOffice had already re-pointed every reference and name for the
  inserted column, so the script does not redo it. `verify_20261127.py` checks that the user's
  edits changed no value of `Spatac-20261126a.xlsx` (with the column shift), and that the
  update changed no value of the baseline but the cells this section changes.
- The Python scripts stay: they are the only way to edit the workbook through LibreOffice
  itself, which recalculates and keeps formats, names and validations; the Rust library cannot
  write a workbook this complex without loss.

### 17.2 Self-review of every build (J8)

The defect files of 2026-11-21 to 2026-11-27 kept returning to the same kinds of defect: text
that does not fit, formats that do not show as set, leftover or misplaced cells, protection, and
regressions of earlier fixes. Two tools now check what a build looks like, and every build runs
them before it is handed over:

- `tools/migrate/layout_lint.py` checks every tab: protection and unlocked cells; a conditional
  font colour overridden by a number format's colour code; a conditional format on empty cells
  only; a discrepancy count without the Alarm format; a number shown as "-0"; and, on every tab
  but the RFS tabs, fit (17.3), word breaks, input rows with a fixed height, and wrapped text in
  hidden columns. It measures each cell by copying it, with its formats, to a scratch tab and
  asking LibreOffice for its optimal width and height. On a fixed-height row, text must fit when
  it renders 10% wider (Excel, 110–120% zoom). A heading in a fixed-height row may not break at
  a hyphen: "Late-/year" was the reported defect, and the narrowest margin missed it, so the
  check fills lines as the programs do.
- `tools/migrate/render.py` prints tabs or ranges to PDF at a chosen zoom, with row and column
  headings. The renders of this build were read page by page; they found the "-0" values, the
  hidden column J making one-line rows two lines tall, and the wrapped end-of-items marker,
  which the check then learned to catch (the first two) or the build fixed (the third).
- `verify_20261127.py` also repeats one check for each earlier defect that a build can bring
  back: top alignment, locked identity columns, "Late-year" wording, Settings last, "RFS total"
  wording, explicit zeros in column S, right-aligned legends, the bold S2 year, the right-aligned
  Adjustments Check, red negatives on Projected_Balances, no leftover labels, the visible
  inflator, the RFS block without green.
- The build sizes columns and rows with the check's own measurements (`needed_width`,
  `needed_height`), so the build and the check cannot disagree.

### 17.3 Row heights and the user's text (J9, J11, J12)

- LibreOffice cannot resize a row by hand on a protected tab (its Protect Sheet options have no
  "format rows", unlike Excel). What works in both programs, protected or not, is an automatic
  row height: the row re-fits whenever a cell in it is edited. So every row with an input cell
  has an automatic height, and the build sets no fixed height on one. The Adjustments input rows
  had a fixed one-line height, which is why E42's note was cut. Fixed heights computed by the
  build were rejected: notes change over time.
- An automatic row is measured with no margin: the program fits it. At some zooms (110%)
  LibreOffice draws a fitted row one line short on screen; no setting in the file fixes it,
  and the row prints and shows whole at 100%. `docs/workbook.md` says so.
- Expenditures column C (Building element) wraps in the item rows, so long descriptions take a
  second line instead of being cut at column D (J9). The end-of-items marker (C250) does not
  wrap.
- Text typed into an input cell, or into a number column ("incl in 18.3.1" in the estimate
  columns), is the user's and is not measured for width (J11).
- The README colour legend's Yellow line (A16) was an unlocked input cell; it is a sample of
  the colour, so it is locked.

### 17.4 Alarm formats (J2, J5)

- The RFS totals check moved to Expenditures T1 with the user's edit; its alarm moves with it,
  and the empty P3 loses its conditional format.
- Projected_Balances C (Expected costs) and I (Diff from RFS_cashflow balance) have the number
  format `#,##0;[RED]-#,##0`, whose `[RED]` overrode the alarm's white text: negative values
  stayed red on red. The alarm on these two columns uses a new cell style, `AlarmNumber`: red
  fill, bold white text and its own number format `#,##0;-#,##0`. In the saved file the alarm's
  `dxf` carries that number format, so Excel and Calc both show white text; outside the alarm
  the columns keep their red negatives.
- Projected_Balances I is rounded to whole dollars, `=ROUND(H−balance;0)`: floating-point
  remainders around 1e-8 showed as a red "-0" in 22 years. Whole dollars are the check's own
  resolution (it flags a difference over 0.5). The rounding was made without asking first; the
  user accepted it but wants such changes proposed beforehand.
- With whole dollars in I, its count I12 is `=SUMPRODUCT(--(I14:I63<>0))`, written like F12
  (was two `COUNTIF`s over ±0.5; the result is the same). F12 and Expenditures T1 keep
  `SUMPRODUCT`: they compute each year's difference inside the formula, which `COUNTIF` cannot.
  The `--` turns TRUE/FALSE into 1/0, which Excel's `SUMPRODUCT` needs. The alarm rule on
  I14:I63 (`ABS(N(I14))>0.5`) is left as it is: on whole dollars it flags the same years.

### 17.5 Expenditures layout (J3)

- Row 7 headings L7:P7 put the year on a line of its own: `=Cfg_BaseYear&CHAR(10)&"Late-year
  estimate"`. "Mid-year" and "Late-year" no longer break at their hyphen; L:P are widened
  together to fit "Approved", the widest word.
- The hidden columns E–H and J no longer wrap: J's adjustment labels, wrapped in a hidden
  column, made one-line item rows two lines tall.
- Widened to fit with the margin: A (Row #), B (item numbers), D (GL accounts), V:AY (the RFS
  totals in rows 3–5, up to 4,265,166).
- Other tabs, widened by the same measurement: Projected_Balances D, E, F, I, M, P, Q, R, V, Z
  (headings that broke inside a word, "RFS_cashflow" among them) and B, R (the summary labels);
  Medium_Term_Cashflow A (the Schedule_of_Investments label was cut); Schedule_of_Investments
  A, B, C, E, F, G, J ("Compounding" was cut); Settings A.

- Row 7 is the AutoFilter's header: each heading's drop-down button covers about 0.45 cm of the
  cell's right side (17 pixels at 100%). The first build ignored it, and the user's screenshot
  showed "RFS chec/k" (I7), "2026 Approve/d" (L7) and "2026 Late-/year/estimate" cut off (O7).
  The layout check now takes 0.5 cm off every AutoFilter heading, and the build widens to fit
  (widths in characters, as `xlsx-diff` reports them): A 5.90 → 8.46, D 10.97 → 11.83, I 7.76 →
  9.84, L:P 10.97 → 13.51, S 12.61 → 12.75.
- A folded column group must fit when unfolded: the hidden past years (V, W, 2024–2025) are
  widened with the rest of the RFS block, and the layout check measures numbers in hidden
  columns.

### 17.6 The RFS tabs (J4, J10)

- RFS_info, RFS_fvalues and RFS_cashflow are protected without a password, every cell locked.
  `spatac-check` requires it. The new-RFS procedure unprotects them to paste and protects them
  again.
- Their description columns (RFS_info C and D, RFS_fvalues B) wrap, and their item rows have an
  automatic height; no fixed heights on a data source. Text that still does not show fully is
  left so (J10). The layout check runs only its protection and format checks there.

### 17.7 `xlsx-diff` (J6)

- Comparing `Spatac-20261126a.xlsx` with the user's edits reported two columns inserted and one
  deleted for the one column inserted, 498 cell changes, and `Exp_NextYear` as
  `$U$7:$R$250`. The 2027 column's formulas use absolute references (`$U8:$AX8`) that the
  insert moved, so its content no longer matched itself and it was paired with nothing. Columns
  are now aligned a second time with the old formulas rewritten through the first alignment,
  until the alignment stops changing; a range whose rows or columns were all deleted reads
  `#REF!` instead of a range whose ends cross. The same comparison now reports one column
  inserted (T) and four cell changes.
- Formats (cell formats, column widths, row heights) are reported by default; `--no-formats`
  leaves them out, and `--formats` is still accepted.

### 17.8 Outputs

- The user renamed the first outputs to `Spatac-20261127.xlsx` and `InvestmentPlan-20261127.xlsx`
  (the first build of a day takes no letter). The second build, with the filter-button widths
  and I12, is `Spatac-20261127a.xlsx`.
- The Investment Plan made to check a build goes to the scratch area, not `_nogit`.

### 17.9 Noted, not changed

- Medium_Term_Cashflow's 2028 adjustment (311) leaves a 5-year investment of 0.44, which shows
  as 0 and matures at 0.51, shown as 1, on Schedule_of_Investments row "Planned 5-year
  investment, bought 2028". A plug of 311.44 would clear it; it is the user's input.
- The Investment Plan's Cover lines A2 and A16 run past column B: whole on screen, cut if
  printed with the print area at B.
