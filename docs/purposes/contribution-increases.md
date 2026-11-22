# Estimating future contribution increases

The current funding plan fixes contributions until the next RFS. After that, the next study
will set new increases, and it helps to know ahead of time roughly what they will need to be.
**Projected_Balances** answers this. It projects the fund on Spatac's plan with the
contribution increases you choose, and shows how close the balance comes to the minimum.

## The inputs (yellow, top left)

| Cell | Meaning |
|---|---|
| B4 | The last year whose contributions follow the current RFS schedule. Later years grow by the increases below. |
| B5–B7 | The increase in each of the first three years after that: the next study's catch-up period. |
| B8 | The annual increase after those three years. It defaults to the current study's subsequent increase. |

## The result

The table at G4:I10 shows the lowest margin over the minimum in the next 30 years, and the year
it occurs, three ways (see [Margins over the minimum](validate-funding.md#margins-over-the-minimum)).
The projected rows show the balance, interest and contributions year by year.

A lowest margin below zero means the increases are not enough. Well above zero means they could
be lower.

## Finding the smallest increase that works

Use **Goal Seek** to find the catch-up increase that brings the lowest margin to zero:

- **Excel:** Data › What-If Analysis › Goal Seek.
- **LibreOffice:** Tools › Goal Seek.

Set the cell **H5** (lowest margin over the minimum, projected) to the value **0** by changing
**B5**. Then copy B5 into B6 and B7 if the three catch-up years should rise equally, and repeat
until B5 stops changing.

Goal Seek searches by trial. If it reports no solution, start B5 from a value nearer the answer,
for example the current study's increase.
