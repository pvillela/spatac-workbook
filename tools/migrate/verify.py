"""Checks a restructured workbook against the workbook it was migrated from.

Usage:

    /opt/libreoffice26.2/program/python tools/migrate/verify.py ORIGINAL.xlsx RESTRUCTURED.xlsx

Both workbooks are recalculated in LibreOffice, then:

1. Key results must match: every computed row of Medium_Term_Cashflow; Expenditures totals per
   year (RFS, current estimates, revised); each investment's interest and maturity values;
   Projected_Balances balances, contributions, interest, adjustments, margins and their 30-year
   lows; and the recomputed RFS balance.
2. No formula on any tab may use a function outside the allowed set (design section 3.3), a
   spill reference or LAMBDA.
3. No formula cell on a Spatac tab may hold an error.

Exits 1 on any failure. The only differences tolerated are listed in EXPECTED, each with its
reason. Anything else is either explained by a phase 4 fix (pass `--expect-changes`, which reports
differences without failing) or it is a migration bug.
"""

import argparse
import re
import sys

from uno_session import Session

TOL = 0.01  # dollars

# Differences the restructure is meant to make, each with its reason. Keyed by the start of the
# difference's description.
EXPECTED = {
    "Revised total 2040": "old fvalues_spatac AC219 held 24,024, a rounded paste of the RFS amount "
    "24,023.75; Expenditures now reads the RFS amount itself (design Q28)",
}

ALLOWED = {
    "ABS", "AGGREGATE", "AND", "COLUMN", "COUNTIFS", "DATEDIF", "FILTER", "IF", "IFERROR",
    "INDEX", "INT", "ISERROR", "ISNA", "LET", "MATCH", "MAX", "MIN", "MOD", "N", "OR", "ROUND",
    "ROW", "ROWS", "SUM", "SUMIFS", "SUMPRODUCT", "TEXT", "TRIM", "VALUE", "XLOOKUP", "XMATCH",
    "YEAR", "TRANSPOSE", "SEQUENCE", "SORT", "UNIQUE", "COUNT", "COUNTA", "COUNTIF", "SUMIF",
    "LEN", "LEFT", "RIGHT", "MID", "CONCATENATE", "DATE", "MONTH", "DAY", "ISBLANK", "NOT",
    "FIND",  # the Expenditures item sort key (design 13.5)
}
SPATAC_TABS = [
    "README", "Settings", "Expenditures", "Medium_Term_Cashflow", "Schedule_of_Investments",
    "Projected_Balances",
]


def col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def colnum(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


class Book:
    def __init__(self, doc):
        self.doc = doc
        self.cache = {}

    def row_of(self, name):
        """The first row a defined name refers to, 1-based."""
        return self.doc.NamedRanges.getByName(name).ReferredCells.RangeAddress.StartRow + 1

    def v(self, sheet, a):
        c = self.doc.Sheets.getByName(sheet).getCellRangeByName(a)
        if c.getError():
            return "#ERR%d" % c.getError()
        t = c.Type.value
        if t == "EMPTY":
            return 0.0
        if t == "TEXT":
            return c.String
        return c.Value


class Checker:
    def __init__(self, old, new):
        self.old, self.new = Book(old), Book(new)
        self.failures = []
        self.expected = []
        self.checked = 0

    def same(self, what, a, b, negate=False):
        self.checked += 1
        if isinstance(a, str) or isinstance(b, str):
            ok = a == b
        else:
            ok = abs((-a if negate else a) - b) <= TOL
        if not ok:
            for prefix, reason in EXPECTED.items():
                if what.startswith(prefix):
                    self.expected.append("%s: original %r, restructured %r — %s" % (what, a, b, reason))
                    return
            self.failures.append("%s: original %r, restructured %r" % (what, a, b))

    def pair(self, what, old_sheet, old_cell, new_sheet, new_cell, negate=False):
        self.same(
            "%s (%s!%s vs %s!%s)" % (what, old_sheet, old_cell, new_sheet, new_cell),
            self.old.v(old_sheet, old_cell),
            self.new.v(new_sheet, new_cell),
            negate,
        )


def check_results(ck):
    # Medium_Term_Cashflow: every computed row, B..G.
    for c in "BCDEFG":
        # Maturities were one row; held and planned are now two.
        ck.same(
            "Medium_Term_Cashflow maturities %s10 vs %s10+%s11" % (c, c, c),
            ck.old.v("Medium_Term_Cashflow", "%s10" % c),
            ck.new.v("Medium_Term_Cashflow", "%s10" % c) + ck.new.v("Medium_Term_Cashflow", "%s11" % c),
        )
    # The restructured rows are found by their defined names, so that moving a row (phase 4)
    # does not break the comparison.
    names = {
        5: ("MTC_OpeningCash", 0),
        8: ("MTC_Contributions", 0),
        9: ("MTC_Expenditures", 0),
        12: ("MTC_NewInvTotal", 0),
        17: ("MTC_NewInv", 4),
        21: ("MTC_ClosingCash", 0),
        24: ("MTC_MinCash", 0),
        28: ("MTC_RFSExpenditures", 0),
        29: ("MTC_RevisedExpenditures", 0),
        30: ("MTC_ExpVariance", 0),
        32: ("MTC_RFSBalance", 0),
        33: ("MTC_ProjectedBalance", 0),
        34: ("MTC_BalanceVariance", 0),
    }
    for r, (name, offset) in names.items():
        nr = ck.new.row_of(name) + offset
        for c in "BCDEFG":
            ck.pair("Medium_Term_Cashflow row %d" % r, "Medium_Term_Cashflow", "%s%d" % (c, r), "Medium_Term_Cashflow", "%s%d" % (c, nr))

    # Expenditures totals by year. Old row 6 = -SUM(column); new row 4 = SUM(column).
    old_current = {2026: "L", 2027: "P", 2028: "Q"}
    for k in range(28):
        year = 2026 + k
        rfs_col = col(colnum("W") + k)
        rev_col = col(colnum("AY") + k)
        old_col = old_current.get(year) or col(colnum("R") + (year - 2029))
        ck.pair("Revised total %d" % year, "fvalues_spatac", "%s6" % old_col, "Expenditures", "%s4" % rev_col, negate=True)
        # RFS per vendor: old row 4 links the vendor total.
        old_rfs = {2026: "H", 2027: "N", 2028: "O"}.get(year) or old_col
        ck.pair("RFS vendor total %d" % year, "fvalues_spatac", "%s4" % old_rfs, "Expenditures", "%s3" % rfs_col, negate=True)
        # The items add up to the vendor's total.
        ck.same(
            "RFS items vs vendor total %d" % year,
            ck.new.v("Expenditures", "%s3" % rfs_col),
            ck.new.v("Expenditures", "%s4" % rfs_col),
        )
    for old_c, new_c in (("L", "N"), ("P", "Q"), ("Q", "T")):
        ck.pair("Current estimate total", "fvalues_spatac", "%s6" % old_c, "Expenditures", "%s4" % new_c, negate=True)

    # Investments: annual interest, maturity value and last-year cash flow.
    for r in range(7, 32):
        for c in "FHIJ":
            ck.pair("Investment row %d" % r, "Schedule_of_Investments", "%s%d" % (c, r), "Schedule_of_Investments", "%s%d" % (c, r))

    # Projected_Balances against the simulator (rows 15..64 -> 13..62) and cashflow_spatac.
    sim = "cashflow_spatac_simulator"
    for i in range(50):
        o, n = 15 + i, 13 + i
        for oc, nc, what in (
            ("A", "A", "year"),
            ("B", "B", "RFS costs"),
            ("D", "L", "interest"),
            ("E", "K", "contributions"),
            ("H", "M", "projected balance"),
            ("M", "E", "RFS balance"),
            ("N", "N", "variance"),
            ("P", "O", "inflator"),
            ("S", "P", "margin over minimum"),
            ("T", "Q", "margin over inflated minimum"),
            ("U", "R", "margin PV"),
        ):
            ck.pair("Projected %s" % what, sim, "%s%d" % (oc, o), "Projected_Balances", "%s%d" % (nc, n))
        old_c = ck.old.v(sim, "C%d" % o)
        ck.same("Projected adjustment (%s!C%d vs Projected_Balances!H%d)" % (sim, o, n), old_c, ck.new.v("Projected_Balances", "H%d" % n))
        for oc, nc, what in (
            ("G", "Z", "recomputed RFS balance"),
            ("C", "X", "recomputed RFS interest"),
            ("D", "Y", "recomputed RFS contributions"),
            ("P", "S", "RFS margin over minimum"),
            ("Q", "T", "RFS margin over inflated minimum"),
            ("R", "U", "RFS margin PV"),
        ):
            ck.pair("RFS check %s" % what, "cashflow_spatac", "%s%d" % (oc, o), "Projected_Balances", "%s%d" % (nc, n))
    for oc, nr in (("S", 5), ("T", 6), ("U", 7)):
        ck.pair("30-year low", sim, "%s13" % oc, "Projected_Balances", "H%d" % nr)
    for oc, nr in (("P", 8), ("Q", 9), ("R", 10)):
        ck.pair("RFS 30-year low", "cashflow_spatac", "%s13" % oc, "Projected_Balances", "H%d" % nr)


def check_formulas(doc):
    problems = []
    fn = re.compile(r"([A-Z][A-Z0-9._]*)\(")
    for sheet in doc.Sheets:
        ranges = sheet.queryContentCells(16)
        for i in range(ranges.Count):
            rng = ranges.getByIndex(i)
            formulas = rng.getFormulaArray()
            addr = rng.RangeAddress
            for dr, row in enumerate(formulas):
                for dc, f in enumerate(row):
                    if not f.startswith("="):
                        continue
                    stripped = re.sub(r'"[^"]*"', '""', f)
                    where = "%s!%s%d" % (sheet.Name, col(addr.StartColumn + dc + 1), addr.StartRow + dr + 1)
                    for name in fn.findall(stripped):
                        base = name.split(".")[-1]
                        if base not in ALLOWED:
                            problems.append("%s uses %s: %s" % (where, name, f))
                    if "#" in stripped.replace("#REF!", "").replace("#N/A", ""):
                        problems.append("%s has a spill reference: %s" % (where, f))
                    if sheet.Name in SPATAC_TABS:
                        cell = sheet.getCellByPosition(addr.StartColumn + dc, addr.StartRow + dr)
                        if cell.getError():
                            problems.append("%s evaluates to error %d: %s" % (where, cell.getError(), f))
    return problems


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("original")
    ap.add_argument("restructured")
    ap.add_argument("--expect-changes", action="store_true", help="report differences without failing")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()
    with Session() as s:
        old = s.open(args.original)
        new = s.open(args.restructured)
        old.calculateAll()
        new.calculateAll()
        ck = Checker(old, new)
        check_results(ck)
        problems = check_formulas(new)
        old.close(True)
        new.close(True)
    print("Key results compared: %d; differences: %d" % (ck.checked, len(ck.failures)))
    for e in ck.expected:
        print("  EXPECTED " + e)
    for f in ck.failures[: (20 if args.quiet else None)]:
        print("  DIFF " + f)
    print("Formula problems: %d" % len(problems))
    for p in problems[:50]:
        print("  " + p)
    if problems or (ck.failures and not args.expect_changes):
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
