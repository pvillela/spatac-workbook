# Validating the RFS funding proposal

An RFS recommends a funding plan: the contribution increases needed to keep the fund above its
minimum balance for the length of the study. Those increases rest on the study's assumptions
about interest, inflation and the planning horizon. Overly cautious assumptions raise
contributions more than necessary. **Projected_Balances** shows what the assumptions produce
and whether the vendor's arithmetic holds.

## The two projections

Projected_Balances has one row per year of the study, 2024 to 2073 for the 2024 study. Years
past the study's 30-year horizon (rows 44–63) are shaded light grey.

- **Columns B–M, the RFS:** the vendor's projection, recomputed in the workbook. Its expected
  costs add up the items of RFS_fvalues for each year. The other inputs are read from
  RFS_cashflow: its interest rates (G7:G8), its contribution increases (K9:K12), its opening
  balance (C11) and its present contribution (C12). Contributions, interest and balances are
  worked out here. The block exists to check the vendor's arithmetic (see below); nothing else
  in the workbook reads it.
- **Columns O–Z, the what-if:** Spatac's plan, from the inputs in P4:P10 (above it). It reads
  the RFS's figures from the RFS tabs, never from the RFS block. Its **Difference from RFS
  balance** compares the two blocks' closing balances.

Both blocks have the same columns and compute each year the same way, which is the vendor's
method:

| Column (RFS, what-if) | |
|---|---|
| B, O Opening balance | The previous year's closing balance. In the first year, the RFS's opening balance (RFS block) or P5 (what-if). The what-if is blank before the current year. |
| C, P Expenditures | The RFS's expected costs, the sum of its items; for the what-if, Spatac's (negative). |
| D, Q Contribution increase | The increase for the year. |
| E, R Contributions | Last year's contributions grown by the increase. The study's first year has the RFS's present contribution. |
| F, S Interest rate | The study's short-term rate for the years set on Settings (B9: 3, so 2024–2026), the long-term rate after. |
| G, T Interest | Opening balance × interest rate. |
| H, U Closing balance | Opening balance + expenditures + contributions + interest. |

## Checking the vendor's arithmetic

Two checks compare the recomputation with RFS_cashflow, year by year. A year that differs by
more than 0.5 shows in red with bold white text, and a count above the block, also red when it
is not 0, gives the number of such years. Both counts should be **0**.

- **Expected costs:** each year's sum of the RFS_fvalues items against the expected costs
  RFS_cashflow states. The count is **Cost years that differ**. A year that differs means the
  vendor's cash flow does not use its own item costs.
- **Closing balance:** the **Diff from RFS_cashflow balance** column, against the balance
  RFS_cashflow states. The count is **Balance years that differ**. A year that differs means
  the vendor's balance does not follow from its own costs, contributions, increases and
  interest rates.

Raise any difference with the vendor. (Expenditures checks the third figure, the vendor's
annual totals on RFS_fvalues, against the same items: see
[Tracking projects](track-projects.md).)

## Margins over the minimum

The RFS states a minimum desired balance (RFS_cashflow C10, $110,000 in the 2024 study) in
study-year dollars. Both projections compare each year's balance with it three ways:

| Columns (RFS, what-if) | Margin |
|---|---|
| J, W | balance minus the minimum as stated; |
| K, X | balance minus the minimum inflated to that year; |
| L, Y | balance in study-year dollars minus the minimum (PV margin). |

A summary above each block gives, for each margin, its lowest value and its first negative
value, each with its year ("none" when the margin never goes negative):

- above the RFS, over the study's first 30 years (2024–2053 for the 2024 study);
- above the what-if, from the current year through the study's last year plus the **Years
  between RFS updates** on Settings (3): 2026–2056 today. That span covers the 30 years of the
  next RFS, so the what-if previews it, before any change to replacement costs. Its heading
  shows the span.

A lowest margin far above zero means the funding plan is more cautious than the minimum
requires. It is worth checking:

- **Interest:** the rates the study assumed (RFS_cashflow G7:G8) against what the fund's
  investments actually earn (Schedule_of_Investments column E).
- **Inflation:** the assumed rates (K7:K8) against recent history.
- **Horizon:** how far out the lowest margin falls. A low point near the end of the study
  depends on cost estimates decades away.

## Spatac's plan

The what-if projects the fund on Spatac's plan rather than the vendor's:

- It starts in the current year (P4) from the **opening balance** in P5. Type it from the
  prior year's financial statements: cash, less payables and accrued liabilities, plus
  investment principal and accrued interest.
- Expenditures are Spatac's: the current year's **Current estimate**, and for later years the
  **adjusted RFS** amounts (the RFS with the Adjustments tab applied) through the end of the
  study's 30-year span. Years after the span use the sum of the RFS items.
- Contribution increases follow the RFS schedule through P6, then the increases in P7:P10. See
  [Estimating future contribution increases](contribution-increases.md). Its first year's
  contributions grow from RFS_cashflow's contribution schedule for the year before.

Its **Difference from RFS balance** then shows how far Spatac's plan departs from the vendor's in each year.
