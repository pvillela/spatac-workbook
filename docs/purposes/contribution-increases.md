# Estimating future contribution increases

The current funding plan fixes contributions until the next RFS. After that, the next study
will set new increases, and it helps to know ahead of time roughly what they will need to be.
**Projected_Balances** answers this. Its what-if projection (columns J–U) projects the fund on
Spatac's plan with the contribution increases you choose, and shows how close the balance comes
to the minimum.

## The inputs (yellow, top left)

| Cell | Meaning |
|---|---|
| B4 | **Current year**: the projection starts this year, from the opening balance in B5. |
| B5 | The opening balance of the current year, from the prior year's financial statements: cash, less payables and accrued liabilities, plus investment principal and accrued interest. |
| B6 | **Revised**: Spatac's expenditures from the Expenditures tab. |
| B7 | **Increases**: contributions grow by the increases below. |
| B8 | The last year whose contributions follow the current RFS schedule. Later years grow by the increases below. |
| B9–B11 | The increase in each of the first three years after that: the next study's catch-up period. |
| B12 | The annual increase after those three years. It defaults to the current study's subsequent increase. |

Column L shows the increase applied in each year, and column M the resulting contributions;
column U gives them per unit per month.

## The result

The table at G4:I10 shows the lowest margin over the minimum in the study's first 30 years, and
the year it occurs, three ways (see [Margins over the minimum](validate-funding.md#margins-over-the-minimum)).
The what-if columns show the balance, interest and contributions year by year.

A lowest margin below zero means the increases are not enough. Well above zero means they could
be lower.

## Finding the smallest increase that works

Use **Goal Seek** to find the catch-up increase that brings the lowest margin to zero:

- **Excel:** Data › What-If Analysis › Goal Seek.
- **LibreOffice:** Tools › Goal Seek.

Set the cell **H5** (lowest what-if margin over the minimum) to the value **0** by changing
**B9**. Then copy B9 into B10 and B11 if the three catch-up years should rise equally, and repeat
until B9 stops changing.

Goal Seek searches by trial. If it reports no solution, start B9 from a value nearer the answer,
for example the current study's increase.
