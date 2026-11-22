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

**Save the workbook before running it.** The tool reads the values the workbook last
calculated. It never changes the source.

## What the plan contains

| Tab | Content |
|---|---|
| Cover | What the plan is, the source file, the date generated and the base year, with lines to record the meeting that approved it and the minutes reference. |
| Medium_Term_Cashflow | The cash section and the RFS-versus-Revised comparison. |
| Schedule_of_Investments | Held and planned investments, with their cash flows by year. |
| Changes_from_RFS | Every item whose Revised amounts differ from the RFS over the six-year horizon: the RFS years and amounts, the revised ones, the net change and the reason. A missing reason shows as **REASON MISSING**. |

Formulas that refer only to the two cash-flow tabs stay formulas, so totals can still be traced.
Formulas that refer to any other tab become their values. Every cell is locked and every tab
protected, without a password.

## What it prints

```
Wrote InvestmentPlan-20261121.xlsx
  1377 formulas kept, 61 replaced by their values
  55 changes from the RFS, 0 without a reason
```

Before sending the plan out, make the last number zero: give every changed item a reason in
Expenditures column U, save, and extract again.

## When it refuses

- "has no defined name …": the workbook is not a restructured Spatac workbook, or a name was
  deleted. See the [reference](../reference/names.md).
- "holds the error …": a cell the plan needs shows an error. Fix it in the workbook, save, and
  run again.
- "exists; pass --force": an InvestmentPlan for that date is already there.
