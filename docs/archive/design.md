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
  - "Vendor total" is now **RFS total**, with legends "RFS total", "Total" and "Variance vs RFS"
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
