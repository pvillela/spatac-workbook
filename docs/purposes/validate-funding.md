# Validating the RFS funding proposal

An RFS recommends a funding plan: the contribution increases needed to keep the fund above its
minimum balance for the length of the study. Those increases rest on the study's assumptions
about interest, inflation and the planning horizon. Overly cautious assumptions raise
contributions more than necessary. **Projected_Balances** shows what the assumptions produce
and whether the vendor's arithmetic holds.

## The two projections

Projected_Balances has one row per year of the study, 2024 to 2073 for the 2024 study.

- **Columns B–H, the RFS:** the vendor's own figures, read from RFS_cashflow: expected costs,
  interest, contributions and balance, and the balance's margins over the minimum. Read only.
- **Columns J–U, the what-if:** the same calculation redone in the workbook, from the choices in
  B4:B12 (yellow, top left). Column Q is its **difference from the RFS balance**.

The what-if computes, for each year, one column at a time. That is the vendor's method.

| Column | |
|---|---|
| J Opening balance | The previous year's closing balance; in the start year, the starting balance (D5). Blank before the start year, and so is the rest of the row. |
| K Expenditures | The Revised or RFS expenditures (negative). |
| L Contribution increase | The increase for the year, if contributions are not following the RFS. |
| M Contributions | The RFS contributions, or last year's grown by the increase. |
| N Interest rate | The study's short-term rate for its first four years, the long-term rate after. |
| O Interest | Opening balance × interest rate. |
| P Closing balance | Opening balance + expenditures + contributions + interest. |

## Re-doing the vendor's arithmetic

Set the what-if to the vendor's assumptions:

| Cell | Set to |
|---|---|
| B4 What-if starts from | **RFS**: the study's first year, from the study's opening balance. |
| B6 Expenditures | **RFS**: the study's expected costs. |
| B7 Contributions | **RFS**: the study's contributions. |

Column Q should then be **zero in every year**. A year that is not means the vendor's balance
does not follow from its own costs, contributions and interest rates. Raise it with the vendor.

To check the vendor's contribution increases as well, set B7 to **Increases**, B8 to the study
year (2024), and B9:B11 to the study's three catch-up increases (RFS_cashflow K9:K11). B12
already follows the study's subsequent increase. Column Q should again be zero.

Set the choices back to **Current year**, **Revised** and **Increases** afterwards.

## Margins over the minimum

The RFS states a minimum desired balance (RFS_cashflow C10, $110,000 in the 2024 study) in
study-year dollars. Both projections compare each year's balance with it three ways:

| Columns (RFS, what-if) | Margin |
|---|---|
| F, P | balance minus the minimum as stated; |
| G, Q | balance minus the minimum inflated to that year; |
| H, R | balance in study-year dollars minus the minimum (PV). |

The table at G4:I10 gives the lowest of each over the study's first 30 years and the year it
occurs.

A lowest margin far above zero means the funding plan is more cautious than the minimum
requires. It is worth checking:

- **Interest:** the rates the study assumed (RFS_cashflow G7:G8) against what the fund's
  investments actually earn (Schedule_of_Investments column E).
- **Inflation:** the assumed rates (K7:K8) against recent history. Projected_Balances E8 records
  13.91% for 2021–2024.
- **Horizon:** how far out the lowest margin falls. A low point near the end of the study
  depends on cost estimates decades away.

## Spatac's plan

With its usual choices (**Current year**, **Revised**, **Increases**), the what-if projects the
fund on Spatac's plan rather than the vendor's:

- It starts in the current year from the **opening balance** in B5. Type it from the prior
  year's financial statements: cash, less payables and accrued liabilities, plus investment
  principal and accrued interest.
- Expenditures are Spatac's: the current year's **Current estimate**, and for later years the
  **adjusted RFS** amounts (the RFS with the Adjustments tab applied) through the end of the
  study's 30-year span. Years before the current year, and years after the span, use the RFS
  amounts.
- Contributions follow the RFS schedule through B8, then grow by the increases in B9:B12. See
  [Estimating future contribution increases](contribution-increases.md).

Column Q then shows how far Spatac's plan departs from the vendor's in each year.
