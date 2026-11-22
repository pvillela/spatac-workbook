# Editing in Excel or LibreOffice

The workbook is an `.xlsx` file and works in current Excel and in LibreOffice Calc 26.2. Either
can be used to edit it. A few rules keep it working in both.

## Rules for formulas

- **Enter formulas normally.** Do not use Ctrl+Shift+Enter.
- **No formulas that spill.** In Excel, a formula such as `=SEQUENCE(5)` or `=FILTER(...)`
  entered in one cell fills the cells below it. LibreOffice 26.2 cannot keep such a formula: it
  freezes it at its size when saved. Functions like `FILTER` are fine inside a formula that
  returns one value, such as `=SUM(FILTER(...))`.
- **No `LAMBDA`,** and no `#` references such as `A1#`. LibreOffice does not support them.
- **No `OFFSET` or `INDIRECT`.** They break when rows move, and the diff tool cannot follow them.
- **No Excel Tables** (Insert › Table). Use defined names instead.
- **Refer to other tabs through defined names** where one exists, so that the reference
  survives inserted rows.

## When two people edit

The workbook is not merged automatically. When someone (Bruce, say) edits a copy while you edit
the master:

1. Ask for the file they started from as well as their edited copy.
2. Compare the two:

   ```
   xlsx-diff Spatac-20261121.xlsx Spatac-20261121-Bruce.xlsx -o bruce.md
   ```

   The report lists every cell they changed, with the old and new formula or value, and every
   comment they added.
3. Make the same changes in your master, by hand.
4. Compare your master before and after, to check you got them all.

Where possible, avoid parallel edits: pass the master back and forth, so that only one person
holds it at a time.

## After saving in a different application

Each application rewrites parts of the file it does not need to: styles, number formats and
column widths. Run `xlsx-diff` between the version before and after the save. Without
`--formats`, the report shows only real changes. If it lists formula or value changes nobody
made, investigate before carrying on.
