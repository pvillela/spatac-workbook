# Estimating future contribution increases

The current funding plan fixes contributions until the next RFS. After that, the next study
will set new increases, and it helps to know ahead of time roughly what they will need to be.
**Projected_Balances** answers this. Its what-if projection (columns O–Z) projects the fund on
Spatac's plan with the contribution increases you choose, and shows how close the balance comes
to the minimum.

## The inputs (yellow, P4:P10, above the what-if)

| Cell | Meaning |
|---|---|
| P4 | The year the projection starts: the current year (Settings B4). Calculated. |
| P5 | The opening balance of the current year, from the prior year's financial statements: cash, less payables and accrued liabilities, plus investment principal and accrued interest. |
| P6 | The last year whose contribution increase follows the current RFS schedule. Set it to 2073 to follow the schedule throughout. |
| P7–P9 | The increase in each of the first three years after that: the next study's catch-up period. |
| P10 | The annual increase after those three years. It defaults to the current study's subsequent increase. |

Each year's contributions are the previous year's grown by that year's increase; the year
before the current one comes from the RFS block. Column Q shows the increase applied in each
year, and column R the resulting contributions; column Z gives them per unit per month.

## The result

The summary at R4:W7 shows, three ways, the lowest margin over the minimum in the study's first
30 years and the first negative one, each with its year (see [Margins over the minimum](validate-funding.md#margins-over-the-minimum)).
The what-if columns show the balance, interest and contributions year by year.

A lowest margin below zero means the increases are not enough. Well above zero means they could
be lower.

## Finding the smallest increase that works

Use **Goal Seek** to find the catch-up increase that brings the lowest margin to zero:

- **Excel:** Data › What-If Analysis › Goal Seek.
- **LibreOffice:** Tools › Goal Seek.

Set the cell **T5** (lowest what-if margin over the minimum) to the value **0** by changing
**P7**. Then copy P7 into P8 and P9 if the three catch-up years should rise equally, and repeat
until P7 stops changing.

Goal Seek searches by trial. If it reports no solution, start P7 from a value nearer the answer,
for example the current study's increase.
