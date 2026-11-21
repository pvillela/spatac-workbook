"""Checks an InvestmentPlan workbook against the Spatac workbook it was extracted from.

Usage:

    /opt/libreoffice26.2/program/python tools/migrate/check_investment_plan.py SPATAC.xlsx PLAN.xlsx

Both are recalculated in LibreOffice. The plan must hold no error anywhere, no formula that
refers to a tab it does not contain, and the same values as the source in every cell it kept on
Medium_Term_Cashflow and Schedule_of_Investments. Exits 1 on any failure.
"""

import re
import sys

from uno_session import Session

KEPT = ("Medium_Term_Cashflow", "Schedule_of_Investments")


def cells(sheet):
    """(row, col) -> (value or text, formula) for every non-empty cell."""
    out = {}
    ranges = sheet.queryContentCells(1 | 2 | 4 | 16)
    for i in range(ranges.Count):
        rng = ranges.getByIndex(i)
        addr = rng.RangeAddress
        data = rng.getDataArray()
        formulas = rng.getFormulaArray()
        for dr, row in enumerate(data):
            for dc, v in enumerate(row):
                r, c = addr.StartRow + dr, addr.StartColumn + dc
                cell = sheet.getCellByPosition(c, r)
                out[(r + 1, c + 1)] = (v, formulas[dr][dc], cell.getError())
    return out


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        return 2
    source, plan = sys.argv[1:]
    failures = []
    with Session() as s:
        src = s.open(source)
        out = s.open(plan)
        src.calculateAll()
        out.calculateAll()
        names = set(out.Sheets.ElementNames)
        print("Plan tabs: %s" % ", ".join(out.Sheets.ElementNames))
        compared = 0
        for name in out.Sheets.ElementNames:
            got = cells(out.Sheets.getByName(name))
            for (r, c), (v, f, err) in got.items():
                if err:
                    failures.append("%s R%dC%d: error %d (%s)" % (name, r, c, err, f))
                if f.startswith("="):
                    for ref in re.findall(r"\$?'?([A-Za-z_][A-Za-z0-9_ ]*)'?\.\$?[A-Z]+\$?\d", f):
                        if ref not in names:
                            failures.append("%s R%dC%d refers to tab %s: %s" % (name, r, c, ref, f))
            if name in KEPT:
                want = cells(src.Sheets.getByName(name))
                for (r, c), (v, f, err) in got.items():
                    compared += 1
                    w = want.get((r, c))
                    if w is None:
                        failures.append("%s R%dC%d not in the source" % (name, r, c))
                        continue
                    a, b = w[0], v
                    if isinstance(a, float) and isinstance(b, float):
                        ok = abs(a - b) <= 0.005
                    else:
                        ok = a == b
                    if not ok:
                        failures.append("%s R%dC%d: source %r, plan %r" % (name, r, c, a, b))
        src.close(True)
        out.close(True)
    print("Cells compared with the source: %d" % compared)
    print("Failures: %d" % len(failures))
    for f in failures[:40]:
        print("  " + f)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
