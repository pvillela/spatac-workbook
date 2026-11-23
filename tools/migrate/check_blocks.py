"""Checks the block boundaries and sorting of a restructured workbook (design 13.5).

Usage:

    /opt/libreoffice26.2/program/python tools/migrate/check_blocks.py RESTRUCTURED.xlsx

The workbook is opened read-only and changed only in memory:

1. Rows are inserted just below the top boundary and just above the ▼ End row of the
   Expenditures items and of the Schedule_of_Investments Held section. Every row-block name must
   grow, and Medium_Term_Cashflow, Projected_Balances, the Expenditures totals and the Planned
   list must not change.
2. On a fresh copy, the Expenditures items are sorted by GL account and by the item sort key.
   ▼ End must stay last and the results must not change. Sorting by Row # must then restore
   every value.
3. On a fresh copy, the January rollover of Medium_Term_Cashflow (design 13.8): the base year goes
   up by one, the first year column is deleted, and a column is inserted left of ▶ End and
   filled from its neighbour. Each remaining year must keep its typed inputs and its looked-up
   figures, no Spatac tab may show an error, and the year check must be clear.
4. On a fresh copy, the Expenditures RFS check (design 13.4): it must be clear on every row; an
   S typed on a row the RFS covers must show "In RFS: clear S", and an S cleared from a Spatac
   row must show "No RFS amounts" or "Not in RFS".

Exits 1 on any failure.
"""

import sys

import uno
from uno_session import Session
from com.sun.star.table import TableSortField


def snapshot(d):
    out = {}
    for n in ("Medium_Term_Cashflow", "Projected_Balances"):
        out[n] = d.Sheets.getByName(n).getCellRangeByName("A1:Z70").getDataArray()
    exp = d.Sheets.getByName("Expenditures")
    out["exp_tot"] = exp.getCellRangeByName("A3:BZ5").getDataArray()
    return out


def close(a, b):
    """Equal, allowing a cent on numbers: a sort changes the order of the additions."""
    for ra, rb in zip(a, b):
        for x, y in zip(ra, rb):
            if isinstance(x, float) and isinstance(y, float):
                if abs(x - y) > 0.01:
                    return False
            elif x != y:
                return False
    return True


def rows_of(d, prefix):
    res = {}
    for n in d.NamedRanges.ElementNames:
        if n.startswith(prefix):
            a = d.NamedRanges.getByName(n).ReferredCells.RangeAddress
            res[n] = a.EndRow - a.StartRow + 1
    return res


def end_row_of(sheet, column):
    for r in range(0, 400):
        if sheet.getCellByPosition(column, r).String.startswith("▼ End"):
            return r + 1
    raise RuntimeError("no ▼ End row on %s" % sheet.Name)


class Checks:
    def __init__(self):
        self.failures = []

    def __call__(self, ok, what):
        print(("ok   " if ok else "FAIL ") + what)
        if not ok:
            self.failures.append(what)


def check_inserts(s, path, check):
    d = s.open(path)
    d.calculateAll()
    exp = d.Sheets.getByName("Expenditures")
    sch = d.Sheets.getByName("Schedule_of_Investments")
    exp_end, held_end = end_row_of(exp, 2), end_row_of(sch, 0)
    held_top = d.NamedRanges.getByName("Inv_Held").ReferredCells.RangeAddress.StartRow + 1
    exp_top = d.NamedRanges.getByName("Exp_Item").ReferredCells.RangeAddress.StartRow + 1
    plan = d.NamedRanges.getByName("Inv_Planned").ReferredCells.RangeAddress
    base = snapshot(d)
    base_plan = sch.getCellRangeByPosition(0, plan.StartRow, 23, plan.EndRow + 2).getDataArray()
    exp0, inv0 = rows_of(d, "Exp_"), rows_of(d, "Inv_")
    for sh in (exp, sch):
        sh.unprotect("")
    # Insert the higher row first, so that the lower row number stays valid.
    exp.Rows.insertByIndex(exp_end - 1, 1)  # just above ▼ End
    exp.Rows.insertByIndex(exp_top, 1)  # just below the header row
    sch.Rows.insertByIndex(held_end - 1, 1)  # just above ▼ End
    sch.Rows.insertByIndex(held_top, 1)  # just below the Held label
    d.calculateAll()
    exp1, inv1 = rows_of(d, "Exp_"), rows_of(d, "Inv_")
    fixed = ("Exp_RFSYears", "Exp_RevisedYears")
    grown = [n for n in exp0 if n not in fixed and exp1[n] != exp0[n] + 2]
    check(not grown, "every Expenditures row-block name grew by 2 %s" % grown)
    for n in ("Inv_Held", "Inv_HeldFlows", "Inv_Name", "Inv_Amount", "Inv_YearFlows"):
        check(inv1[n] == inv0[n] + 2, "%s grew by 2" % n)
    check(inv1["Inv_Planned"] == inv0["Inv_Planned"], "Inv_Planned kept its size")
    after = snapshot(d)
    for n in ("Medium_Term_Cashflow", "Projected_Balances"):
        check(after[n] == base[n], "%s unchanged" % n)
    check(after["exp_tot"] == base["exp_tot"], "Expenditures rows 3-5 unchanged")
    moved = sch.getCellRangeByPosition(0, plan.StartRow + 2, 23, plan.EndRow + 4).getDataArray()
    check(moved == base_plan, "Planned list and totals unchanged (moved down 2 rows)")
    d.close(True)


def check_sorts(s, path, check):
    d = s.open(path)
    d.calculateAll()
    exp = d.Sheets.getByName("Expenditures")
    exp.unprotect("")
    key_col = d.NamedRanges.getByName("Exp_SortKey").ReferredCells.RangeAddress.StartColumn
    hdr = d.NamedRanges.getByName("Exp_Item").ReferredCells.RangeAddress.StartRow
    end = end_row_of(exp, 2)
    rng = exp.getCellRangeByPosition(0, hdr, key_col, end - 1)
    keys = {r[1]: r[key_col] for r in rng.getDataArray()[1:-1]}
    for k in ("1", "2.1", "10.1", "3.10a", "5.3.1", "5.3.1a", "15.5a", "18.3.1"):
        print("   sort key %-8s %s" % (k, keys.get(k)))
    check(keys["2.1"] < keys["10.1"] < keys["18.3.1"], "sort keys order 2.1, 10.1, 18.3.1")
    check(keys["5.3.1"] < keys["5.3.1a"], "sort keys order 5.3.1 before 5.3.1a")
    before = rng.getDataArray()
    base = snapshot(d)

    def sort_by(column):
        field = TableSortField()
        field.Field = column
        field.IsAscending = True
        field.IsCaseSensitive = False
        desc = rng.createSortDescriptor()
        for p in desc:
            if p.Name == "SortFields":
                p.Value = uno.Any("[]com.sun.star.table.TableSortField", (field,))
            if p.Name == "ContainsHeader":
                p.Value = True
        rng.sort(desc)
        d.calculateAll()

    for column, what in ((3, "GL account"), (key_col, "item sort key")):
        sort_by(column)
        check(rng.getDataArray()[-1][2].startswith("▼ End"), "sorted by %s: ▼ End is still last" % what)
        now = snapshot(d)
        for n in ("Medium_Term_Cashflow", "exp_tot"):
            check(close(now[n], base[n]), "sorted by %s: %s unchanged" % (what, n))
    sort_by(0)
    check(rng.getDataArray() == before, "sorting by Row # restores every value")
    d.close(True)


def by_year(sheet, years_row, rows, years):
    """{(row, year): value} for the given rows, read under each year heading."""
    heads = sheet.getCellRangeByPosition(0, years_row - 1, 30, years_row - 1).getDataArray()[0]
    out = {}
    for c, y in enumerate(heads):
        if isinstance(y, float) and int(y) in years:
            for r in rows:
                out[(r, int(y))] = sheet.getCellByPosition(c, r - 1).getValue()
    return out


def check_rollover(s, path, check):
    d = s.open(path)
    d.calculateAll()
    names = d.NamedRanges
    mtc = d.Sheets.getByName("Medium_Term_Cashflow")
    base = int(names.getByName("Cfg_BaseYear").ReferredCells.getCellByPosition(0, 0).Value)
    years_row = names.getByName("MTC_Years").ReferredCells.RangeAddress.StartRow + 1
    grid = names.getByName("MTC_NewInv").ReferredCells.RangeAddress
    rows_of = lambda n: names.getByName(n).ReferredCells.RangeAddress.StartRow + 1
    inputs = [rows_of("MTC_Payables"), rows_of("MTC_CarryOver"), rows_of("MTC_Adjustment")]
    inputs += list(range(grid.StartRow + 1, grid.EndRow + 1))  # the 1- to 4-year rows
    looked_up = [rows_of(n) for n in ("MTC_Contributions", "MTC_Maturities", "MTC_RFSExpenditures")]
    kept = set(range(base + 1, base + 6))
    before = by_year(mtc, years_row, inputs + looked_up, kept)
    opening_next = by_year(mtc, years_row, [rows_of("MTC_OpeningCash")], {base + 1})

    for name in ("Medium_Term_Cashflow", "Settings"):
        d.Sheets.getByName(name).unprotect("")
    names.getByName("Cfg_BaseYear").ReferredCells.getCellByPosition(0, 0).setValue(base + 1)
    d.calculateAll()
    check(mtc.getCellRangeByName("A2").String != "", "the year check warns until the columns follow")
    first = names.getByName("MTC_Years").ReferredCells.RangeAddress.StartColumn
    mtc.Columns.removeByIndex(first, 1)
    end = [c for c in range(0, 30) if mtc.getCellByPosition(c, years_row - 1).String.startswith("▶ End")][0]
    mtc.Columns.insertByIndex(end, 1)
    source = mtc.getCellRangeByPosition(end - 1, 0, end - 1, 60).RangeAddress
    mtc.copyRange(mtc.getCellByPosition(end, 0).CellAddress, source)
    mtc.getCellByPosition(end, years_row - 1).setValue(base + 6)
    opening = mtc.getCellByPosition(first, rows_of("MTC_OpeningCash") - 1)
    opening.setValue(opening_next[(rows_of("MTC_OpeningCash"), base + 1)])
    d.calculateAll()

    years = names.getByName("MTC_Years").ReferredCells.getDataArray()[0]
    check(list(years[:-1]) == [float(y) for y in range(base + 1, base + 7)], "years now %d-%d" % (base + 1, base + 6))
    warning = mtc.getCellRangeByName("A2").String
    check(warning == "", "the year check is clear %r" % warning)
    after = by_year(mtc, years_row, inputs + looked_up, kept)
    moved = [k for k in before if abs(before[k] - after.get(k, float("nan"))) > 0.01]
    check(not moved, "each year %d-%d kept its inputs and looked-up figures %s" % (base + 1, base + 5, moved[:5]))
    errors = []
    for name in ("Expenditures", "Medium_Term_Cashflow", "Schedule_of_Investments", "Projected_Balances"):
        sh = d.Sheets.getByName(name)
        found = sh.queryFormulaCells(4)  # com.sun.star.sheet.FormulaResult.ERROR
        for i in range(found.Count):
            errors.append("%s %s" % (name, found.getByIndex(i).AbsoluteName))
    check(not errors, "no errors on the Spatac tabs %s" % errors[:5])
    grid = names.getByName("MTC_NewInv").ReferredCells.getDataArray()
    nonzero = sum(1 for row in grid for v in row if isinstance(v, float) and abs(v) > 0.005)
    planned = names.getByName("Inv_Planned").ReferredCells.getDataArray()
    listed = sum(1 for row in planned if row[0] != "")
    check(listed == nonzero, "the Planned list has the %d grid amounts (%d listed)" % (nonzero, listed))
    d.close(True)


def check_rfs_flags(s, path, check):
    d = s.open(path)
    d.calculateAll()
    exp = d.Sheets.getByName("Expenditures")
    end = end_row_of(exp, 2)
    top = d.NamedRanges.getByName("Exp_Item").ReferredCells.RangeAddress.StartRow + 1
    rows = exp.getCellRangeByPosition(0, top, 8, end - 2).getDataArray()
    flagged = [top + 1 + k for k, row in enumerate(rows) if row[8] != ""]
    check(not flagged, "the RFS check is clear on every row %s" % flagged[:5])
    spatac = next(top + 1 + k for k, row in enumerate(rows) if row[5] == "S" and row[3] != "")
    covered = next(top + 1 + k for k, row in enumerate(rows) if row[5] == "" and row[3] != "")
    exp.unprotect("")
    exp.getCellRangeByName("F%d" % covered).setString("S")
    exp.getCellRangeByName("F%d" % spatac).setString("")
    d.calculateAll()
    now = exp.getCellRangeByName("I%d" % covered).getString()
    check(now == "In RFS: clear S", "an S on row %d, which the RFS covers, is flagged %r" % (covered, now))
    now = exp.getCellRangeByName("I%d" % spatac).getString()
    check(
        now in ("No RFS amounts", "Not in RFS"),
        "Spatac row %d without its S is flagged %r" % (spatac, now),
    )
    d.close(True)


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    check = Checks()
    with Session() as s:
        check_inserts(s, sys.argv[1], check)
        check_sorts(s, sys.argv[1], check)
        check_rollover(s, sys.argv[1], check)
        check_rfs_flags(s, sys.argv[1], check)
    print("Failures: %d" % len(check.failures))
    return 1 if check.failures else 0


if __name__ == "__main__":
    sys.exit(main())
