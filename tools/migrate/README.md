# Workbook migration

One-off tooling that restructured the Spatac workbook (design: `docs/archive/design.md`,
phases 3 and 4). Kept for reference and so the restructure can be re-run against a newer
source workbook.

Both scripts drive headless LibreOffice 26.2 and must be run with its bundled Python:

```
/opt/libreoffice26.2/program/python tools/migrate/migrate.py _nogit/Spatac-20261121.xlsx _nogit/Spatac-20261121-restructured.xlsx
/opt/libreoffice26.2/program/python tools/migrate/verify.py _nogit/Spatac-20261121.xlsx _nogit/Spatac-20261121-restructured.xlsx
```

- `migrate.py` builds the restructured workbook from the source, which it opens read-only. It
  prints what it changed in the data. It always applies the two phase 4 bug fixes; there is no
  way to build the old bugs back in:
  - investments paying interest at maturity compound over their own term, not always 5 years;
  - the Medium_Term_Cashflow closing balance counts the adjustment once; the carry-over row
    moves below it.
- `verify.py` recalculates both workbooks and compares every key result. It fails on any
  difference not listed, with its reason, in `EXPECTED`: one restructure change and the 13
  results the phase 4 fixes change, each listed by its cells. It also checks that every formula uses
  allowed functions only and that no formula on a Spatac tab evaluates to an error.

## The Adjustments restructure (design section 14)

`restructure.py` does not rebuild the workbook: it edits the current build in place, adding the
Adjustments tab and reshaping Expenditures, Medium_Term_Cashflow, Projected_Balances, Settings and
README. It reads the user's six adjustments from the mock-up.

```
/opt/libreoffice26.2/program/python tools/migrate/restructure.py _nogit/Spatac-20261123c.xlsx _nogit/Spatac-20261123c-restructured.xlsx OUT.xlsx --notes NOTES.txt
/opt/libreoffice26.2/program/python tools/migrate/verify_restructure.py _nogit/Spatac-20261123c.xlsx OUT.xlsx
```

- `restructure.py` writes OUT once and refuses an existing file. `--notes` lists every working
  note whose column letters it translated, for review.
- `verify_restructure.py` requires numeric parity with the source, after patching in the one
  deliberate fix (item 3.9's 2028 amount, design 14.4); exercises each Adjustments check and one
  valid move; and checks that the what-if still reproduces the RFS.

The scripts open a copy of each source, made in a throwaway folder, so LibreOffice never leaves a
lock file beside a source in `_nogit`.

Neither script writes to the source workbook. Each starts LibreOffice with a throwaway profile, so
it does not touch your own LibreOffice settings or a running instance.
