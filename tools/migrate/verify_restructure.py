"""Checks a restructured workbook against the build it came from (design section 14).

Usage:

    /opt/libreoffice26.2/program/python tools/migrate/verify_restructure.py OLD.xlsx NEW.xlsx

OLD is the source of restructure.py (Spatac-20261123c.xlsx). Numeric parity is required, apart
from one deliberate fix (design 14.4, G32): item 3.9's 2028 amount carried two years' inflation
instead of three. OLD is patched in memory with the fixed amount, and then every key result must
match:

- Medium_Term_Cashflow, every row, allowing for the inserted row 24;
- Projected_Balances, every year and the lowest margins, in their new columns (design 15.4);
- Expenditures: the current year's Current estimate per item, and each later year of the old
  Revised block against the adjusted block;
- no error values on any Spatac tab, and no flagged Adjustments rows;
- the recomputed RFS block matches RFS_cashflow, and its check catches a changed balance;
- the what-if follows the RFS schedule when told to, and the interest years follow Settings.

Both files are opened read-only.
"""

import sys

import uno  # noqa: F401
from uno_session import Session

TOL = 0.005
SPATAC_TABS = ["README", "Expenditures", "Adjustments", "Medium_Term_Cashflow",
               "Schedule_of_Investments", "Projected_Balances", "Settings"]
GARAGE_RAMP_2028 = 229500 * 1.02 ** 3


def block(doc, sheet, rng):
    return [list(r) for r in doc.Sheets.getByName(sheet).getCellRangeByName(rng).getDataArray()]


def same(a, b):
    if isinstance(a, float) and isinstance(b, float):
        return abs(a - b) <= TOL
    if a in ("", 0.0) and b in ("", 0.0):
        return True
    return a == b


def col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class Checks:
    def __init__(self):
        self.failures = []
        self.count = 0

    def eq(self, what, a, b):
        self.count += 1
        if not same(a, b):
            self.failures.append("%s: old %r, new %r" % (what, a, b))

    def true(self, what, ok):
        self.count += 1
        if not ok:
            self.failures.append(what)


def find_row(doc, sheet, column, text, first=1, last=400):
    data = block(doc, sheet, "%s%d:%s%d" % (column, first, column, last))
    rows = [first + i for i, (v,) in enumerate(data) if str(v).strip() == text]
    if len(rows) != 1:
        raise RuntimeError("%s!%s: %r found %d times" % (sheet, column, text, len(rows)))
    return rows[0]


def colnum(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def self_check_projected(c, new):
    """The recomputed RFS block and the simpler what-if (design 15.1-15.3), in memory."""
    pb = new.Sheets.getByName("Projected_Balances")
    cf = new.Sheets.getByName("RFS_cashflow")
    settings = new.Sheets.getByName("Settings")
    get = lambda rng: block(new, "Projected_Balances", rng)

    # The recomputation matches RFS_cashflow in every year.
    diffs = [v for (v,) in get("I15:I64") if v == "" or abs(v) > TOL]
    c.true("Projected_Balances: recomputed RFS differs from RFS_cashflow: %s" % diffs[:3], not diffs)
    c.eq("Projected_Balances I13 (years that differ)", 0.0, get("I13:I13")[0][0])
    c.eq("Settings A9", "Years of short-term interest in the RFS (the long-term rate applies after them)",
         block(new, "Settings", "A9:A9")[0][0])
    c.eq("Cfg_ShortInterestYears", 3.0,
         new.NamedRanges.getByName("Cfg_ShortInterestYears").ReferredCells.getValue())

    # A changed vendor balance is caught, in that year only.
    g20 = cf.getCellRangeByName("G20")
    kept = g20.getValue()
    g20.setValue(kept + 1000)
    new.calculateAll()
    c.eq("Projected_Balances I13 after RFS_cashflow G20 + 1000", 1.0, get("I13:I13")[0][0])
    c.eq("Projected_Balances I20 after RFS_cashflow G20 + 1000", -1000.0, get("I20:I20")[0][0])
    g20.setValue(kept)

    # Through the chosen year the what-if's increase is the RFS's; through 2073 its contributions
    # are the RFS's throughout.
    rows = get("A15:AB64")
    through = get("P6:P6")[0][0]
    for r in rows:
        if r[14] != "" and r[0] <= through and r[0] > rows[0][0]:
            c.eq("Projected_Balances Q for %d is the RFS's increase" % r[0], r[3], r[16])
    pb.getCellRangeByName("P6").setValue(2073)
    new.calculateAll()
    for r in get("A15:AB64"):
        if r[14] != "":
            c.eq("Projected_Balances R for %d with the RFS schedule throughout" % r[0], r[4], r[17])
    pb.getCellRangeByName("P6").setValue(through)

    # Interest: the short-term rate for the first Cfg_ShortInterestYears years.
    g8 = cf.getCellRangeByName("G8")
    long_rate = g8.getValue()
    g8.setValue(0.05)
    settings.getCellRangeByName("B9").setValue(1)
    new.calculateAll()
    rates = [r[5] for r in get("A15:F16")]
    c.true("Projected_Balances F15:F16 with one short-term year: %s" % rates,
           abs(rates[0] - cf.getCellRangeByName("G7").getValue()) < 1e-9 and abs(rates[1] - 0.05) < 1e-9)
    g8.setValue(long_rate)
    settings.getCellRangeByName("B9").setValue(3)
    new.calculateAll()


def main():
    old_path, new_path = sys.argv[1], sys.argv[2]
    c = Checks()
    with Session() as s:
        old = s.open_copy(old_path)
        new = s.open_copy(new_path)
        for d in (old, new):
            for sheet in d.Sheets:
                if sheet.isProtected():
                    sheet.unprotect("")

        # The deliberate fix, applied to OLD: 3.9's 2028 Late-year estimate (old S).
        r39 = find_row(old, "Expenditures", "B", "3.9")
        old.Sheets.getByName("Expenditures").getCellRangeByName("S%d" % r39).setValue(GARAGE_RAMP_2028)
        old.calculateAll()
        new.calculateAll()

        # Medium_Term_Cashflow: rows 1-23 in place; 24 onward one row down.
        o = block(old, "Medium_Term_Cashflow", "A1:G36")
        n = block(new, "Medium_Term_Cashflow", "A1:G37")
        for r in range(36):
            nr = r if r < 23 else r + 1
            for k in range(1, 7):
                if r == 0:
                    continue  # the title carries the preparation date
                c.eq("Medium_Term_Cashflow %s%d" % (col(k + 1), r + 1), o[r][k], n[nr][k])
        c.true("Medium_Term_Cashflow row 24 is blank", all(v == "" for v in n[23][:7]))
        c.eq("Medium_Term_Cashflow A25", "Minimum closing cash balance", n[24][0])

        # Projected_Balances: every year and the lowest margins, old column -> new column. The
        # what-if's increase (old L) was blank while contributions followed the RFS; it now shows
        # the RFS's increase, so only its other years are compared.
        moved = {"A": "A", "B": "C", "C": "G", "D": "E", "E": "H", "F": "J", "G": "K", "H": "L",
                 "J": "O", "K": "P", "L": "Q", "M": "R", "N": "S", "O": "T", "P": "U", "Q": "V",
                 "R": "W", "S": "X", "T": "Y", "U": "Z"}
        o = block(old, "Projected_Balances", "A15:U64")
        n = block(new, "Projected_Balances", "A15:AB64")
        for r in range(50):
            for oc, nc in moved.items():
                a, b = o[r][colnum(oc) - 1], n[r][colnum(nc) - 1]
                if oc == "L" and a == "":
                    continue
                c.eq("Projected_Balances %s%d -> %s%d" % (oc, r + 15, nc, r + 15), a, b)
        # The lowest margins: old H5:I10 (what-if rows 5-7, RFS rows 8-10) -> the what-if summary
        # U5:V7 and the RFS summary E5:F7.
        o = block(old, "Projected_Balances", "H5:I10")
        for r in range(6):
            at, row = ("U", r + 5) if r < 3 else ("E", r + 2)
            got = block(new, "Projected_Balances", "%s%d:%s%d" % (at, row, col(colnum(at) + 1), row))[0]
            for k in range(2):
                c.eq("Projected_Balances %s%d -> %s%d" % ("HI"[k], r + 5, col(colnum(at) + k), row),
                     o[r][k], got[k])
        # The first negative margins, worked out here from the year table's first 30 years.
        years = block(new, "Projected_Balances", "A15:AB44")
        for at, margins in (("G", "JKL"), ("W", "WXY")):
            for i, m in enumerate(margins):
                k = colnum(m) - 1
                neg = [(r[0], r[k]) for r in years if r[k] != "" and r[k] < 0]
                want = [neg[0][1], neg[0][0]] if neg else ["", "none"]
                got = block(new, "Projected_Balances", "%s%d:%s%d" % (at, 5 + i, col(colnum(at) + 1), 5 + i))[0]
                c.eq("Projected_Balances %s%d first negative %s" % (at, 5 + i, m), want[0], got[0])
                c.eq("Projected_Balances %s%d year of first negative %s" % (col(colnum(at) + 1), 5 + i, m),
                     want[1], got[1])

        # Expenditures, item by item (same rows): current year; later years of Revised.
        end = find_row(new, "Expenditures", "C", "▼ End of items: insert new rows above this line")
        o = block(old, "Expenditures", "A8:BZ%d" % (end - 1))
        n = block(new, "Expenditures", "A8:AX%d" % (end - 1))
        o_years = [int(v) for v in block(old, "Expenditures", "AY6:BZ6")[0]]
        n_years = [int(v) for v in block(new, "Expenditures", "U6:AX6")[0]]
        base = o_years[0]
        for i in range(len(o)):
            row = i + 8
            c.eq("Expenditures row %d item" % row, o[i][1], n[i][1])
            c.eq("Expenditures row %d Current estimate" % row, o[i][13], n[i][15])  # N -> P
            c.eq("Expenditures row %d notes kept" % row, bool(o[i][21]), bool(n[i][16]))  # V -> Q
            for k, y in enumerate(o_years):
                if y == base:
                    continue
                c.eq("Expenditures row %d %d" % (row, y), o[i][50 + k], n[i][20 + n_years.index(y)])
            c.eq("Expenditures row %d next year" % row, o[i][50 + 1], n[i][18])  # S

        # No error values; nothing flagged on Adjustments.
        for name in SPATAC_TABS:
            sheet = new.Sheets.getByName(name)
            cur = sheet.createCursor()
            cur.gotoEndOfUsedArea(False)
            a = cur.RangeAddress
            rng = sheet.getCellRangeByPosition(0, 0, a.EndColumn, a.EndRow)
            errors = rng.queryFormulaCells(4)  # com.sun.star.sheet.FormulaResult.ERROR
            c.true("%s: error values in %s" % (name, errors.RangeAddressesAsString),
                   errors.Count == 0)
        checks = [v for (v,) in block(new, "Adjustments", "H1:H80") if v not in ("", "Check", "OK")]
        c.true("Adjustments: flagged rows %s" % checks, not checks)
        c.eq("Adjustments A2", "", block(new, "Adjustments", "A2:A2")[0][0])
        oks = [v for (v,) in block(new, "Adjustments", "H39:H44")]
        c.true("Adjustments: the user's six rows show OK: %s" % oks, oks == ["OK"] * 6)

        # Blank versus deliberate zero (design 13.4): 3.9 has no RFS amount and no adjustment in
        # 2027, so it is blank; its 2025 amount was moved away entirely, so it is 0.
        r39 = find_row(new, "Expenditures", "B", "3.9")
        y = [int(v) for v in block(new, "Expenditures", "U6:AX6")[0]]
        row = block(new, "Expenditures", "A%d:AX%d" % (r39, r39))[0]
        c.true("Expenditures 3.9 2027 is blank: %r" % row[20 + y.index(2027)], row[20 + y.index(2027)] == "")
        c.true("Expenditures 3.9 next year is blank: %r" % row[18], row[18] == "")
        c.true("Expenditures 3.9 2025 is a zero: %r" % row[20 + y.index(2025)], row[20 + y.index(2025)] == 0.0)
        # Variances: more spending than the RFS is negative.
        c2028 = col(21 + y.index(2028))
        v2028 = block(new, "Expenditures", "%s5:%s5" % (c2028, c2028))[0][0]
        c.true("Expenditures 2028 variance is negative: %r" % v2028, v2028 < 0)

        # Adjustments behaviour, in memory, in the last block: each check, and one valid move.
        adj = new.Sheets.getByName("Adjustments")
        r20 = find_row(new, "Expenditures", "B", "20.1")
        label = block(new, "Expenditures", "J%d:J%d" % (r20, r20))[0][0]
        years = [int(v) for v in block(new, "Expenditures", "U6:AX6")[0]]
        y30, y31 = (20 + years.index(2030), 20 + years.index(2031))
        before = block(new, "Expenditures", "A%d:AX%d" % (r20, r20))[0]
        rows = [
            (label, 2028, 2029, 1e9, "x", "More than the item holds in 2028"),
            (label, 2027, 2027, 1, "x", "Year from = Year to"),
            ("Not an item", 2027, 2028, 1, "x", "Item not in Expenditures"),
            (label, "", "", 5, "", "Missing entry"),
            (label, 2030, 2031, 1000, "x", "OK"),
        ]
        for i, (item, frm, to, amount, reason, _) in enumerate(rows):
            r = 55 + i
            for column, v in zip("ABCDE", (item, frm, to, amount, reason)):
                cell = adj.getCellRangeByName("%s%d" % (column, r))
                if isinstance(v, (int, float)):
                    cell.setValue(v)
                elif v:
                    cell.setString(v)
        new.calculateAll()
        for i, row in enumerate(rows):
            got = block(new, "Adjustments", "H%d:H%d" % (55 + i, 55 + i))[0][0]
            c.eq("Adjustments H%d" % (55 + i), row[5], got)
        c.true("Adjustments A2 counts 4 problems",
               block(new, "Adjustments", "A2:A2")[0][0].startswith("Check: 4 row(s)"))
        after = block(new, "Expenditures", "A%d:AX%d" % (r20, r20))[0]
        c.eq("Expenditures 20.1 2030 after the move", before[y30] - 1000, after[y30])
        c.eq("Expenditures 20.1 2031 after the move", before[y31] + 1000 * 1.02, after[y31])
        adj.getCellRangeByName("A55:E59").clearContents(1 | 2 | 4 | 16)
        new.calculateAll()

        self_check_projected(c, new)

        old.close(True)
        new.close(True)

    print("Checks: %d" % c.count)
    print("Failures: %d" % len(c.failures))
    for f in c.failures[:60]:
        print("- " + f)
    return 1 if c.failures else 0


if __name__ == "__main__":
    sys.exit(main())
