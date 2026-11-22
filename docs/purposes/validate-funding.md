# Validating the RFS funding proposal

An RFS recommends a funding plan: the contribution increases needed to keep the fund above its
minimum balance for the length of the study. Those increases rest on the study's assumptions
about interest, inflation and the planning horizon. Overly cautious assumptions raise
contributions more than necessary. **Projected_Balances** shows what the assumptions produce
and whether the vendor's arithmetic holds.

## Re-doing the vendor's arithmetic

The check columns on Projected_Balances (X to AA, folded by default) rebuild the study's cash
flow from its own parameters:

- **X**: interest on the previous year's balance, at the short-term rate for the study's first
  four years and the long-term rate after;
- **Y**: contributions, growing at the study's first-year, second-year and third-year increases
  and then at its subsequent increase;
- **Z**: the balance from those;
- **AA**: flags any year where Z differs from the vendor's balance by more than a dollar.

A flagged year means the vendor's figures do not follow from the parameters it states.
Raise it with the vendor.

## Margins over the minimum

The RFS states a minimum desired balance (RFS_cashflow C10, $110,000 in the 2024 study) in
study-year dollars. Projected_Balances compares each year's balance with it three ways:

| Columns (RFS, projected) | Margin |
|---|---|
| S, P | balance minus the minimum as stated; |
| T, Q | balance minus the minimum inflated to that year; |
| U, R | balance in study-year dollars minus the minimum (PV). |

The table at G4:I10 gives the lowest of each over the next 30 years and the year it occurs.

A lowest margin far above zero means the funding plan is more cautious than the minimum
requires. It is worth checking:

- **Interest:** the rates the study assumed (RFS_cashflow G7:G8) against what the fund's
  investments actually earn (Schedule_of_Investments column E).
- **Inflation:** the assumed rates (K7:K8) against recent history. Projected_Balances E4 records
  13.91% for 2021–2024.
- **Horizon:** how far out the lowest margin falls. A low point near the end of the study
  depends on cost estimates decades away.

## Spatac's changes

The projected columns (H to R) apply Spatac's changes to the study: for each year, Revised minus
RFS from Expenditures (column H), plus the actual expenditures and the actual closing balance
for years already finished (columns F and G, yellow). The projected balance in column M is what
the fund looks like on Spatac's plan rather than the vendor's.
