# xlsx-diff

`xlsx-diff` compares two versions of a workbook and writes what changed as Markdown. It works on
any `.xlsx` workbook, not only the Spatac one.

```
xlsx-diff OLD.xlsx NEW.xlsx [-o REPORT.md] [--values] [--formats]
```

Without `-o`, the report goes to the terminal. Markdown reads well as plain text and renders as
a formatted page in VS Code, on GitHub, or in any Markdown viewer.

## What it reports

**Tabs.** Tabs added, deleted and renamed, and a changed tab order.

- Tabs are identified, not listed cell by cell.
- A renamed tab is recognised by its contents being mostly the same as the deleted tab's.

**Defined names.** Names added, removed, or pointing somewhere else.

**For each tab in both versions:**

- **Structure:**
  - rows and columns inserted, deleted or moved;
  - merged cells;
  - conditional formats and data validation;
  - sheet protection, and cells newly locked or unlocked;
  - hidden rows, columns and tabs;
  - frozen panes.
- **Cell changes:** every cell whose formula, constant or comment changed, with its address in
  both versions and the old and new content. A change made the same way to neighbouring cells
  (a formula filled down or across, say) is listed once, with the range and a count.

**Formulas are compared as written, after allowing for moved rows.** When a row is inserted,
every formula below it, and every formula pointing past it, has its references shifted by the
application. `xlsx-diff` follows the rows and columns between the two versions and rewrites the
old formulas accordingly. A formula that merely moved, or whose references only shifted, is
**not** reported.

## Options

- `--values`: also list cells whose formula is unchanged but whose result changed, in a
  separate, folded section. Without it, one changed input that ripples through a hundred
  totals is reported once, as the input.
- `--formats`: also list changed cell formats, column widths and row heights. Widths and
  heights are compared with a tolerance, because each application rounds them differently on
  saving.
- `-o FILE`: write the report to `FILE`.

## Reading a report

The report starts with a summary: tabs added, deleted, renamed or reordered, and which tabs
changed. Then comes one section per changed tab:

- **Structure:** for example "2 rows inserted (new numbering): 22–23" or "Row 19 moved to row
  23". Row numbers follow the version named.
- **Cell changes:** a table with the cell in the new version, the cell in the old version (when
  different), the kind of change (changed, added, removed, comment added and so on), and the
  content before and after.

A typical review of a month's edits shows a handful of inputs and comments. A long list of
changed formulas nobody intended is the signal to look closer.
