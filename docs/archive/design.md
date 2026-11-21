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

### 3.4 Naming (Q30)

- Defined names are workbook-scoped, with a per-tab prefix plus PascalCase, e.g.
  `MTC_ClosingCash`, `Inv_Due`, `Exp_RFS`, `RFS_Fv_Costs`, `Cfg_BaseYear`.
- A defined name may cover a 2-D block (rows × year columns) with a companion header
  name, e.g. `Exp_RFS` and `Exp_RFSYears`.
- CLI tools refer only to defined names, never to cell addresses.

### 3.5 Inputs and protection (Q6)

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
    `RFS_info`. Spatac rows are exempt.
  - Another flags `RFS_info` items with no Expenditures row.
- The item list is typed, not formula-generated, so inputs stay attached to their rows
  (Q35).
- Green conditional-format row highlighting is removed.
- Migration of today's `=MAX(I,J)` / `=J` formulas in column L: wherever their value
  differs from the fallback rule, it becomes a typed Latest estimate. Results are
  preserved.
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
  Fix, decided by the user: closing balance `=SUM(B5:B12)+B19`.

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
