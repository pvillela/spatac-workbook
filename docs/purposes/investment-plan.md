# The Investment Plan

The Reserve Fund Investment Policy (s.3.2) makes the Medium-Term Cash Flow schedule the
corporation's Investment Plan. The plan:

- starts from the latest RFS;
- adds the Board's changes to project timing, cost and priorities, and the reasons for them;
- matches investment maturities plus the cash contingency, in aggregate, to the expenditures
  expected over a five-year horizon;
- ladders maturities so that something matures at least every year, normally with terms of
  five years or less.

This chapter is the written description of the schedule that s.3.2 asks to be kept with the
policy: its structure, its sources and how it is updated. Export it to PDF from the printer
icon at the top of the page.

## Structure

The plan is the **Medium_Term_Cashflow** tab, with **Schedule_of_Investments** behind it. It
covers six years, the current year and five more, in columns B to G.

| Row | Content | Source |
|---|---|---|
| 5 | Opening cash balance | First year: typed in, the actual cash at the end of last year. Later years: the previous year's closing balance. |
| 6 | Payables and accrued liabilities | Typed in. |
| 8 | Planned contributions | The RFS funding plan (RFS_cashflow). |
| 9 | Revised expenditures | Expenditures, Revised block: the RFS amounts as changed by the Board. |
| 10 | Held investments maturing, plus interest | Schedule_of_Investments, held investments. |
| 11 | Planned investments maturing | The planned investments in rows 14–18 of earlier years, compounded at the assumed rate. |
| 13 | Planned new investments | The total of rows 14–18, as an outflow. |
| 14–17 | 1- to 4-year investments to buy | Typed in. |
| 18 | 5-year investments to buy | Calculated: whatever cash is left above the minimum, the carry-over and the adjustment. |
| 20 | Closing cash balance | Rows 5–13. By construction, the minimum plus the carry-over plus the adjustment. |
| 22 | Additional cash carry-over | Typed in: cash to hold above the minimum. |
| 23 | Adjustment to minimum cash balance | Typed in: a correction to the cash the plan holds, such as a rounding difference with the bank. |
| 24 | Minimum closing cash balance | Settings: the Board's minimum target (s.3.5), inflated from the year it is stated in at the RFS short-term inflation rate. |
| 28–30 | RFS versus Revised expenditures | The Board's changes, year by year. |
| 32–34 | RFS versus projected fund balance | Projected_Balances, for context. Not part of the extracted plan. |

**Schedule_of_Investments** has one row per investment:

- **Held** (rows 7–46, between the Held label and the grey ▼ End row): each investment actually held. Type its name, due date, purchase date,
  book value, rate and compounding (**A**: interest paid annually; **C**: interest paid at
  maturity).
- **Planned** (rows 49–78): filled automatically, one row for each non-zero amount in
  Medium_Term_Cashflow rows 14–18.
- **Year columns** (N onward): each investment's cash flows. An annual-interest investment pays
  its interest in each year after purchase and before maturity. In the maturity year it pays
  back its value plus that year's interest.

## Checking the plan against the policy

- **Liquidity:** closing cash (row 20) never falls below the minimum (row 24), less any
  adjustment.
- **Laddering:** Schedule_of_Investments shows something maturing in every year of the horizon.
- **Terms:** no term (Schedule_of_Investments column F) exceeds five years without the broker's
  advice and a recorded reason.
- **Changes:** every change from the RFS has a reason (see below).

## Updating the plan

The policy (s.3.2) names three occasions for a revision.

**After a new RFS and funding plan is adopted.** Follow
[A new Reserve Fund Study](../procedures/new-rfs.md), then review the planned investments.

**Before any investment or reinvestment the approved plan does not already cover.**

1. Update the held investments.
2. Adjust the amounts in rows 14–17 for the year of purchase. The 5-year row rebalances itself.
3. Extract the plan and send it for approval.

When an investment is actually bought, type it into Held and set its amount in rows 14–18 to
zero. If it was a 5-year investment, change the 1- to 4-year amounts so that the 5-year row
comes out at zero.

**In mid to late January,** with the prior year's closing figures and the Board's November
decisions for the new year. Follow [January rollover](../procedures/rollover.md).

## The extracted plan

`spatac-investment-plan` writes the plan as a workbook of its own. It is what goes to the Board,
the Treasurer and the broker, and it is kept with the approving minutes. It contains:

- a **cover** with the source file, the date and lines for the approving meeting;
- **Medium_Term_Cashflow**, without rows 32–34;
- **Schedule_of_Investments**.

The reasons for the Board's changes from the RFS are not part of the plan: the cover says they
are described in the Spatac workbooks for the current and previous years (Expenditures
column U).

See [spatac-investment-plan](../tools/investment-plan.md).
