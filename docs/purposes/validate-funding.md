# Validating the RFS funding proposal

An RFS recommends a funding plan: the contribution increases needed to keep the fund above its
minimum balance for the length of the study. Those increases rest on the study's assumptions
about interest, inflation and the planning horizon. Overly cautious assumptions raise
contributions more than necessary. **Projected_Balances** shows what the assumptions produce
and whether the vendor's arithmetic holds.

## The two projections

Projected_Balances has one row per year of the study, 2024 to 2073 for the 2024 study. Years
past the study's 30-year horizon (rows 44–63) are shaded light grey.

- **Columns B–M, the RFS:** the vendor's projection, recomputed in the workbook. Only the
  study's inputs are read from RFS_cashflow: its expected costs (column B), its interest rates
  (G7:G8), its contribution increases (K9:K12), its opening balance (C11) and its present
  contribution (C12). Contributions, interest and balances are worked out here. Column I is the
  **difference from RFS_cashflow's balance** (see below).
- **Columns O–Z, the what-if:** Spatac's plan, from the inputs in P4:P10 (above it). Column V
  is its **difference from the RFS balance**.

Both blocks have the same columns and compute each year the same way, which is the vendor's
method:

| Column (RFS, what-if) | |
|---|---|
| B, O Opening balance | The previous year's closing balance. In the first year, the RFS's opening balance (RFS block) or P5 (what-if). The what-if is blank before the current year. |
| C, P Expenditures | The RFS's expected costs; for the what-if, Spatac's (negative). |
| D, Q Contribution increase | The increase for the year. |
| E, R Contributions | Last year's contributions grown by the increase. The study's first year has the RFS's present contribution. |
| F, S Interest rate | The study's short-term rate for the years set on Settings (B9: 3, so 2024–2026), the long-term rate after. |
| G, T Interest | Opening balance × interest rate. |
| H, U Closing balance | Opening balance + expenditures + contributions + interest. |

## Checking the vendor's arithmetic

Column I compares the recomputed balance with the balance RFS_cashflow states, year by year.
It should be **zero in every year**. A year that differs by more than 0.5 is shaded orange, and
I12, above the column, counts those years in red. A year that differs means the vendor's
balance does not follow from its own costs, contributions, increases and interest rates. Raise
it with the vendor.

## Margins over the minimum

The RFS states a minimum desired balance (RFS_cashflow C10, $110,000 in the 2024 study) in
study-year dollars. Both projections compare each year's balance with it three ways:

| Columns (RFS, what-if) | Margin |
|---|---|
| J, W | balance minus the minimum as stated; |
| K, X | balance minus the minimum inflated to that year; |
| L, Y | balance in study-year dollars minus the minimum (PV margin). |

A summary above each block (B4:G7 for the RFS, R4:W7 for the what-if) gives, for each margin
over the study's first 30 years, its lowest value and its first negative value, each with its
year ("none" when the margin never goes negative).

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
  study's 30-year span. Years after the span use the RFS amounts.
- Contribution increases follow the RFS schedule through P6, then the increases in P7:P10. See
  [Estimating future contribution increases](contribution-increases.md).

Column V then shows how far Spatac's plan departs from the vendor's in each year.
