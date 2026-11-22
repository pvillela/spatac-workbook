# Defined names

Every name is workbook-wide. The prefix says which tab it points at. Do not rename or delete a
name: formulas and the command-line tools depend on them.

## Inserting rows

A name grows when a row is inserted **inside** it. Each list of rows on Expenditures and the
Held section of Schedule_of_Investments is framed so that this happens wherever you insert:

- **At the top**, the frame is the row above the first entry: the header row on Expenditures
  (row 7), the "Held" label on Schedule_of_Investments (row 6). It is part of the names, so a
  row inserted just below it is inside them.
- **At the bottom**, the frame is the grey **▼ End** row, which is also part of the names.
  Insert new rows **above** it. A row inserted below ▼ End is outside the names.

The frame rows hold no amounts, so totals and lookups are unaffected. **Do not delete the header
row, the Held label or a ▼ End row**: the names would then stop at the first or last entry
again, and an insert at that edge would fall outside them.

## Settings (`Cfg_`)

| Name | Cell | Meaning |
|---|---|---|
| `Cfg_BaseYear` | B4 | The current year. |
| `Cfg_MinTarget` | B5 | The minimum target cash balance (policy s.3.5). |
| `Cfg_MinTargetYear` | B6 | The year whose dollars the minimum target is stated in. |
| `Cfg_NewInvRate` | B7 | The interest rate assumed for planned investments. |

## Expenditures (`Exp_`)

Each column name runs from the header row (7) through the ▼ End row (250 in the November 2026
workbook).

| Name | Column | Name | Column |
|---|---|---|---|
| `Exp_RowNumber` | A | `Exp_Approved` | J |
| `Exp_Item` | B | `Exp_Actual` | K |
| `Exp_Element` | C | `Exp_MidEst` | L |
| `Exp_GL` | D | `Exp_LateEst` | M |
| `Exp_GLName` | E | `Exp_Current` | N |
| `Exp_SpatacRow` | F | `Exp_MidEst1`, `Exp_LateEst1`, `Exp_Current1` | O, P, Q |
| `Exp_Key` | G | `Exp_MidEst2`, `Exp_LateEst2`, `Exp_Current2` | R, S, T |
| `Exp_Key2` | H | `Exp_Reason` | U |
| `Exp_Check` | I | `Exp_Notes` | V |
| | | `Exp_SortKey` | CA (folded): the item number in a form that sorts correctly |

| Name | Range | Meaning |
|---|---|---|
| `Exp_RFS` | W7:AX250 | RFS amount per item and year. |
| `Exp_RFSYears` | W6:AX6 | The years of `Exp_RFS`. |
| `Exp_Revised` | AY7:BZ250 | Revised amount per item and year. |
| `Exp_RevisedYears` | AY6:BZ6 | The years of `Exp_Revised`. |

## Medium_Term_Cashflow (`MTC_`)

Each covers columns B–G of its row.

| Name | Row | Name | Row |
|---|---|---|---|
| `MTC_Years` | 3 | `MTC_ClosingCash` | 20 |
| `MTC_OpeningCash` | 5 | `MTC_CarryOver` | 22 |
| `MTC_Payables` | 6 | `MTC_Adjustment` | 23 |
| `MTC_Contributions` | 8 | `MTC_MinCash` | 24 |
| `MTC_Expenditures` | 9 | `MTC_RFSExpenditures` | 28 |
| `MTC_Maturities` | 10 | `MTC_RevisedExpenditures` | 29 |
| `MTC_PlannedMaturities` | 11 | `MTC_ExpVariance` | 30 |
| `MTC_NewInvTotal` | 13 | `MTC_RFSBalance` | 32 |
| `MTC_NewInv` | 14–18 | `MTC_ProjectedBalance` | 33 |
| | | `MTC_BalanceVariance` | 34 |

Three more names cover areas rather than a single row:

- `MTC_Title` is A1.
- `MTC_CashSection` is A3:G24.
- `MTC_ComparisonSection` is A28:G30.

The extraction tool keeps these.

## Schedule_of_Investments (`Inv_`)

| Name | Range | Meaning |
|---|---|---|
| `Inv_Name` … `Inv_DueYear` | columns A–L, rows 6–78 | Each investment field, held and planned. |
| `Inv_Held` | A6:X47 | The held investments, from the Held label to ▼ End. |
| `Inv_Planned` | A49:X78 | The planned investments. |
| `Inv_Years` | N5:X5 | The years of the cash-flow columns. |
| `Inv_YearFlows` | N6:X78 | Cash flows, held and planned. |
| `Inv_HeldFlows` | N6:X47 | Cash flows, held only. |
| `Inv_PlannedFlows` | N49:X78 | Cash flows, planned only. |
| `Inv_Totals` | N80:X80 | Total cash flows per year. |
| `Inv_Title` | A1 | The title. |

## Projected_Balances (`PB_`)

| Name | Range | Meaning |
|---|---|---|
| `PB_Start` | B4 | Where the what-if starts: **Current year** or **RFS**. |
| `PB_StartYear` | D4 | The year it starts in. |
| `PB_OpeningBalance` | B5 | The opening balance of the current year. |
| `PB_ExpBasis`, `PB_ContribBasis` | B6, B7 | The what-if's expenditures (**Revised** or **RFS**) and contributions (**Increases** or **RFS**). |
| `PB_ContribFixedThrough` | B8 | Last year of contributions per the RFS schedule. |
| `PB_CatchUp` | B9:B11 | Increases in the three years after it. |
| `PB_SubsequentIncrease` | B12 | The increase after those. |
| `PB_Lowest` | G5:I10 | The 30-year lows. |
| `PB_Year`; `PB_RFSCost` … `PB_RFSMarginPV` | columns A; B–H, rows 15–64 | The years; the RFS baseline. |
| `PB_Cost`, `PB_Increase`, `PB_Contributions`, `PB_Interest`, `PB_Balance`, `PB_Variance`, `PB_Margin`, `PB_MarginInflated`, `PB_MarginPV`, `PB_PerUnitMonth` | columns J–S, rows 15–64 | The what-if. `PB_Variance` is the difference from the RFS balance. |
| `PB_Inflator` | column U, rows 15–64 | The inflation factor from the study year. |

## RFS tabs (`RFS_`)

| Name | Range |
|---|---|
| `RFS_Info_Key`, `RFS_Info_Check` | RFS_info J1:J400, K1:K400 |
| `RFS_Fv_Key` | RFS_fvalues BC1:BC400 |
| `RFS_Fv_Values` | RFS_fvalues C1:AZ400 |
| `RFS_Fv_Years` | RFS_fvalues C9:AZ9 |
| `RFS_Fv_Totals` | RFS_fvalues C10:AZ10 |
| `RFS_Cf_Year`, `RFS_Cf_Cost`, `RFS_Cf_Interest`, `RFS_Cf_Contrib`, `RFS_Cf_Balance` | RFS_cashflow A, B, C, D, G (rows 1–200) |
| `RFS_StudyYear` | RFS_cashflow A15 |
| `RFS_Units`, `RFS_MinBalance`, `RFS_OpeningBalance`, `RFS_PresentContribution` | RFS_cashflow C6, C10, C11, C12 |
| `RFS_ShortInterest`, `RFS_LongInterest` | RFS_cashflow G7, G8 |
| `RFS_ShortInflation`, `RFS_LongInflation` | RFS_cashflow K7, K8 |
| `RFS_CatchUpIncreases`, `RFS_SubsequentIncrease` | RFS_cashflow K9:K11, K12 |
