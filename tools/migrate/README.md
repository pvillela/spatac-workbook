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
  prints what it changed in the data. Without options it reproduces the source's results,
  bugs included (phase 3). The phase 4 fixes are options, and the workbook to use is built with
  both:

  ```
  /opt/libreoffice26.2/program/python tools/migrate/migrate.py --fix compounding --fix closing _nogit/Spatac-20261121.xlsx _nogit/Spatac-20261121-restructured.xlsx
  ```

  - `compounding`: investments paying interest at maturity compound over their own term, not
    always 5 years.
  - `closing`: the Medium_Term_Cashflow closing balance counts the adjustment once; the
    carry-over row moves below it.
- `verify.py` recalculates both workbooks and compares every key result. It fails on any
  difference not listed, with its reason, in `EXPECTED`; pass `--expect-changes` to list the
  differences the phase 4 fixes make without failing. It also checks that every formula uses
  allowed functions only and that no formula on a Spatac tab evaluates to an error.

Neither script writes to the source workbook. Each starts LibreOffice with a throwaway profile, so
it does not touch your own LibreOffice settings or a running instance.
