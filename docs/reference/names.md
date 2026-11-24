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
| `Cfg_ShortInflationYears` | B8 | How many years from the study year the RFS applies its short-term inflation rate; the long-term rate applies after them. |

## Expenditures (`Exp_`)

Each column name runs from the header row (7) through the ▼ End row (250 in the November 2026
workbook).

| Name | Column | Name | Column |
|---|---|---|---|
| `Exp_RowNumber` | A | `Exp_AdjLabel` | J |
| `Exp_Item` | B | `Exp_Approved` | K |
| `Exp_Element` | C | `Exp_Actual` | L |
| `Exp_GL` | D | `Exp_MidEst` | M |
| `Exp_GLName` | E | `Exp_LateEst` | N |
| `Exp_SpatacRow` | F | `Exp_Current` | O |
| `Exp_Key` | G | `Exp_Notes` | P |
| `Exp_Key2` | H | `Exp_NextYear` | R |
| `Exp_Check` | I | | |

| Name | Range | Meaning |
|---|---|---|
| `Exp_Adjusted` | T7:AW250 | Adjusted RFS amount per item and year: the RFS amount with the Adjustments tab applied. |
| `Exp_AdjustedYears` | T6:AW6 | The years of `Exp_Adjusted`: the RFS study year and the 29 after it. |

## Adjustments (`Adj_`)

Each column name runs from the first block's header row (6) to the last block's closing row
(65), so a row inserted inside any block is inside them.

| Name | Column | Meaning |
|---|---|---|
| `Adj_Item` | A | The item, as its Expenditures adjustment label. |
| `Adj_From`, `Adj_To` | B, C | Year from and Year to, or **none**. |
| `Adj_Amount` | D | The amount typed. |
| `Adj_Reason` | E | The reason for the change. |
| `Adj_FromAmount`, `Adj_ToAmount` | F, G | What leaves Year from, in its dollars; what arrives in Year to, in its dollars. |
| `Adj_Check` | H | The row's check. |
| `Adj_YearChoices` | J6:J36 (hidden) | The choices for Year from and Year to: **none** and the RFS's 30 years. |

## Medium_Term_Cashflow (`MTC_`)

Each covers columns B–H of its row: the six year columns and the grey **▶ End** column. Insert a new year column just left of ▶ End, so that every name grows to include it; never delete ▶ End.

| Name | Row | Name | Row |
|---|---|---|---|
| `MTC_Years` | 3 | `MTC_ClosingCash` | 20 |
| `MTC_OpeningCash` | 5 | `MTC_CarryOver` | 22 |
| `MTC_Payables` | 6 | `MTC_Adjustment` | 23 |
| `MTC_Contributions` | 8 | `MTC_MinCash` | 25 |
| `MTC_Expenditures` | 9 | `MTC_RFSExpenditures` | 29 |
| `MTC_Maturities` | 10 | `MTC_RevisedExpenditures` | 30 |
| `MTC_PlannedMaturities` | 11 | `MTC_ExpVariance` | 31 |
| `MTC_NewInvTotal` | 13 | `MTC_RFSBalance` | 33 |
| `MTC_NewInv` | 14–18 | `MTC_ProjectedBalance` | 34 |
| | | `MTC_BalanceVariance` | 35 |

Three more names cover areas rather than a single row:

- `MTC_Title` is A1. A2 is a check that the first year column is the base year and that the
  years run in steps of one.
- `MTC_CashSection` is A3:H25.
- `MTC_ComparisonSection` is A29:H31.

The extraction tool keeps these.

## Schedule_of_Investments (`Inv_`)

| Name | Range | Meaning |
|---|---|---|
| `Inv_Name` … `Inv_DueYear` | columns A–L, rows 6–78 | Each investment field, held and planned: one name per column, `Inv_` and a short form of its heading (`Inv_Name`, `Inv_Due`, `Inv_Purchased`, … `Inv_DueYear`). Exception: Book value (D) is `Inv_Amount`. |
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
| `PB_StartBalance` | D5 | The balance it starts from: the RFS opening balance, or B5. |
| `PB_OpeningBalance` | B5 | The opening balance of the current year. |
| `PB_ExpBasis`, `PB_ContribBasis` | B6, B7 | The what-if's expenditures (**Revised** or **RFS**) and contributions (**Increases** or **RFS**). |
| `PB_ContribFixedThrough` | B8 | Last year of contributions per the RFS schedule. |
| `PB_CatchUp` | B9:B11 | Increases in the three years after it. |
| `PB_SubsequentIncrease` | B12 | The increase after those. |
| `PB_Lowest` | G5:I10 | The 30-year lows. |
| `PB_Year`; `PB_RFSCost` … `PB_RFSMarginPV` | columns A; B–H, rows 15–64 | The years; the RFS baseline: one name per column, `PB_RFS` and a short form of its heading (`PB_RFSCost`, `PB_RFSInterest`, … `PB_RFSMarginPV`). |
| `PB_Opening`, `PB_Cost`, `PB_Increase`, `PB_Contributions`, `PB_Rate`, `PB_Interest`, `PB_Balance`, `PB_Variance`, `PB_Margin`, `PB_MarginInflated`, `PB_MarginPV`, `PB_PerUnitMonth` | columns J–U, rows 15–64 | The what-if. `PB_Variance` is the difference from the RFS balance. |
| `PB_Inflator` | column W, rows 15–64 | The inflation factor from the study year: the short-term rate for `Cfg_ShortInflationYears` years, the long-term rate after. The Adjustments tab inflates and deflates with it. |

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
