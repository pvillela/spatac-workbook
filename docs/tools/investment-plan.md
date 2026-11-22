# spatac-investment-plan

`spatac-investment-plan` extracts the [Investment Plan](../purposes/investment-plan.md) from the
Spatac workbook as a workbook of its own. That workbook is what goes to the Board for approval
and to the Treasurer and the broker.

```
spatac-investment-plan Spatac-20261121.xlsx
```

This writes `InvestmentPlan-20261121.xlsx` beside the source, with the date taken from the source's
name. Options:

- `-o FILE`: write somewhere else, or when the source's name has no date.
- `--force`: replace an existing output file.
- `--date YYYY-MM-DD`: the date to show as "Generated" on the cover. Without it, the cover shows
  today's date in the computer's own time zone.

**Save the workbook before running it.** The tool reads the values the workbook last
calculated. It never changes the source.

## What the plan contains

| Tab | Content |
|---|---|
| Cover | What the plan is, the source file, the date generated and the base year, with lines to record the meeting that approved it and the minutes reference. It says that the reasons for changes from the RFS are described in the Spatac workbooks for the current and previous years. |
| Medium_Term_Cashflow | The cash section and the RFS-versus-Revised comparison. |
| Schedule_of_Investments | Held and planned investments, with their cash flows by year. |

Formulas that refer only to the two cash-flow tabs stay formulas, so totals can still be traced.
Formulas that refer to any other tab become their values. Every cell is locked and every tab
protected, without a password.

## What it prints

```
Wrote InvestmentPlan-20261121.xlsx
  1377 formulas kept, 61 replaced by their values
```

## When it refuses

- "has no defined name …": the workbook is not a restructured Spatac workbook, or a name was
  deleted. See the [reference](../reference/names.md).
- "holds the error …": a cell the plan needs shows an error. Fix it in the workbook, save, and
  run again.
- "exists; pass --force": an InvestmentPlan for that date is already there.
