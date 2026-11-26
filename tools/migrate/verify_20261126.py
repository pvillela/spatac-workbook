"""Checks update_20261126.py's output against its source (design-history.md section 16).

Usage:

    /opt/libreoffice26.2/program/python tools/migrate/verify_20261126.py \
        SOURCE.xlsx EDITS.xlsx NEW.xlsx

SOURCE is the baseline (_nogit/Spatac-20261125h-PV.xlsx); EDITS is the user's edited copy
(_nogit/Spatac-20261125h-PV-edits.xlsx), for the headings and labels; NEW is the output. All are
opened read-only. The checks:

- every value on every tab equals SOURCE's, allowing for the rows inserted on Settings and
  README, apart from the cells section 16 changes on purpose;
- the what-if summary, recomputed here from the what-if's margins over its new span, and again
  after the span is changed;
- no error values on any Spatac tab; no IFERROR or ISNA left where section 16 removed them;
- each new check is 0 today and catches a changed RFS item, vendor total and RFS_cashflow cost;
- the what-if does not depend on the RFS block;
- the new formats, headings, labels, README text and protection.
"""

import re
import sys
import zipfile

import uno  # noqa: F401
from uno_session import Session
from verify_restructure import Checks, block, col, colnum

TOL = 0.005
SPATAC_TABS = ["README", "Expenditures", "Adjustments", "Medium_Term_Cashflow",
               "Schedule_of_Investments", "Projected_Balances", "Settings"]
SAME_TABS = ["Expenditures", "Adjustments", "Medium_Term_Cashflow", "Schedule_of_Investments",
             "RFS_info", "RFS_fvalues", "RFS_cashflow"]
CHANGED = {
    "Expenditures": {"O3", "P3", "Q3"},
    "Projected_Balances": {"A2", "B12", "E12", "F12", "H12", "R4"}
    | {"%s13" % col(k) for k in range(2, 14)}
    | {"%s%d" % (c, r) for c in "TUVW" for r in (5, 6, 7)},
}


def same(a, b):
    if isinstance(a, float) and isinstance(b, float):
        return abs(a - b) <= TOL
    return a == b


def used(sheet):
    cur = sheet.createCursor()
    cur.gotoEndOfUsedArea(False)
    a = cur.RangeAddress
    return a.EndRow + 1, a.EndColumn + 1


def compare(c, old, new, tab, row_map=lambda r: r, skip=()):
    o, n = old.Sheets.getByName(tab), new.Sheets.getByName(tab)
    rows, cols = used(o)
    od = o.getCellRangeByPosition(0, 0, cols - 1, rows - 1).getDataArray()
    for r in range(rows):
        nr = row_map(r + 1)
        nd = n.getCellRangeByPosition(0, nr - 1, cols - 1, nr - 1).getDataArray()[0]
        for k in range(cols):
            a = "%s%d" % (col(k + 1), nr)
            if a in skip:
                continue
            c.count += 1
            if not same(od[r][k], nd[k]):
                c.failures.append("%s!%s: old %r, new %r" % (tab, a, od[r][k], nd[k]))


def no_errors(c, doc):
    for tab in SPATAC_TABS:
        sheet = doc.Sheets.getByName(tab)
        rows, cols = used(sheet)
        bad = []
        for r in range(rows):
            for k in range(cols):
                cell = sheet.getCellByPosition(k, r)
                if cell.getError():
                    bad.append("%s%d" % (col(k + 1), r + 1))
        c.true("%s: error values in %s" % (tab, bad[:5]), not bad)


def summary(doc):
    """The what-if summary, from the margins (W, X, Y) over the span the settings give."""
    names = doc.NamedRanges
    value = lambda n: names.getByName(n).ReferredCells.getValue()
    first = value("Cfg_BaseYear")
    last = value("RFS_StudyYear") + 29 + value("Cfg_YearsBetweenRFS")
    rows = block(doc, "Projected_Balances", "A14:Z63")
    out = []
    for m in (22, 23, 24):  # W, X, Y
        pts = [(r[0], r[m]) for r in rows if first <= r[0] <= last and r[m] != ""]
        low = min(v for _, v in pts)
        low_year = min(y for y, v in pts if v == low)
        negs = [(y, v) for y, v in pts if v < 0]
        out.append([low, low_year] + (list(negs[0][::-1]) if negs else ["", "none"]))
    return first, last, out


def check_summary(c, doc, what):
    first, last, want = summary(doc)
    got = block(doc, "Projected_Balances", "T5:W7")
    for i, (w, g) in enumerate(zip(want, got)):
        for j, (a, b) in enumerate(zip(w, g)):
            c.eq("%s: Projected_Balances %s%d" % (what, "TUVW"[j], 5 + i), a, b)
    c.eq("%s: Projected_Balances R4" % what, "What-if, %d–%d" % (first, last),
         block(doc, "Projected_Balances", "R4:R4")[0][0])


def alarm_at(doc, sheet, address):
    """Whether a conditional format at `address` uses a red fill with bold white text (an
    .xlsx round trip renames the Alarm style)."""
    styles = doc.StyleFamilies.getByName("CellStyles")
    cf = sheet.getCellRangeByName(address).ConditionalFormat
    for i in range(cf.Count):
        st = styles.getByName(cf.getByIndex(i).StyleName)
        if st.CellBackColor == 0xC00000 and st.CharColor == 0xFFFFFF and st.CharWeight == 150.0:
            return True
    return False


def formula_text(sheet, rng):
    return " ".join(f for row in sheet.getCellRangeByName(rng).getFormulaArray() for f in row)


def run(c, old, edits, new, new_path):
    for d in (old, edits, new):
        d.calculateAll()

    # Values.
    for tab in SAME_TABS:
        compare(c, old, new, tab, skip=CHANGED.get(tab, ()))
    compare(c, old, new, "Projected_Balances", skip=CHANGED["Projected_Balances"])
    compare(c, old, new, "Settings", row_map=lambda r: r if r < 10 else r + 2, skip={"A6"})
    compare(c, old, new, "README", row_map=lambda r: r if r < 4 else r + 1,
            skip={"A%d" % r for r in range(1, 60)})
    no_errors(c, new)
    check_summary(c, new, "today")

    exp = new.Sheets.getByName("Expenditures")
    pb = new.Sheets.getByName("Projected_Balances")
    mtc = new.Sheets.getByName("Medium_Term_Cashflow")
    fv = new.Sheets.getByName("RFS_fvalues")
    cf = new.Sheets.getByName("RFS_cashflow")
    get = lambda sheet, a: sheet.getCellRangeByName(a).getDataArray()[0][0]

    # The checks are 0 today, and each catches what it is for.
    for sheet, a in ((exp, "P3"), (pb, "F12"), (pb, "I12")):
        c.eq("%s!%s today" % (sheet.Name, a), 0.0, get(sheet, a))

    def tamper(sheet, a, delta, want):
        cell = sheet.getCellRangeByName(a)
        kept = cell.getValue()
        cell.setValue(kept + delta)
        new.calculateAll()
        for s2, a2, v in want:
            c.eq("after %s!%s + %g: %s!%s" % (sheet.Name, a, delta, s2.Name, a2), v, get(s2, a2))
        cell.setValue(kept)
        new.calculateAll()

    # An item that no longer adds up to the vendor's total, in 2024 (RFS_fvalues C22 is 2.5's).
    tamper(fv, "C22", 1000, [(exp, "P3", 1.0), (pb, "F12", 1.0), (exp, "U3", 597850.0),
                             (pb, "C14", -597850.0)])
    # A vendor total that no longer adds up: Expenditures sees it, Projected_Balances does not.
    tamper(fv, "D10", 1000, [(exp, "P3", 1.0), (pb, "F12", 0.0)])
    # An RFS_cashflow cost that differs from the items, in 2025.
    tamper(cf, "B16", 1000, [(pb, "F12", 1.0), (exp, "P3", 0.0)])
    c.eq("RFS_cashflow A16", 2025.0, get(cf, "A16"))

    # The summary follows the setting.
    years = new.NamedRanges.getByName("Cfg_YearsBetweenRFS").ReferredCells
    years.setValue(5)
    new.calculateAll()
    check_summary(c, new, "with 5 years between RFS updates")
    years.setValue(3)
    new.calculateAll()

    # Formulas that no longer hide errors.
    for sheet, rng in ((pb, "C14:C63"), (pb, "I14:I63"), (pb, "P14:S63"), (exp, "U3:AX3"),
                       (new.Sheets.getByName("Adjustments"), "G1:G200")):
        text = formula_text(sheet, rng)
        c.true("%s!%s: IFERROR or ISNA left" % (sheet.Name, rng),
               not re.search(r"IFERROR|ISNA|IFNA", text))
    c.true("Projected_Balances P14:S63 reads the RFS block",
           not re.search(r"\$[B-M]\d", formula_text(pb, "P14:S63")))

    # The loud formats.
    for sheet, a in ((exp, "U3"), (exp, "AX3"), (exp, "P3"), (pb, "C14"), (pb, "C63"),
                     (pb, "F12"), (pb, "I12"), (pb, "I14"), (mtc, "B14"), (mtc, "G18")):
        c.true("%s!%s has no Alarm format" % (sheet.Name, a), alarm_at(new, sheet, a))
    c.true("Projected_Balances B20 is still green", pb.getCellRangeByName("B20").CellBackColor != 0xE2EFDA)
    c.true("Projected_Balances P20 is no longer green", pb.getCellRangeByName("P20").CellBackColor == 0xE2EFDA)
    c.true("Projected_Balances B50 is not light grey", pb.getCellRangeByName("B50").CellBackColor == 0xEDEDED)
    c.true("Projected_Balances AB hidden", pb.Columns.getByName("AB").IsVisible)
    with zipfile.ZipFile(new_path) as z:
        sheets = [n for n in z.namelist() if n.startswith("xl/worksheets/sheet")]
        grouped = [n for n in sheets if re.search(r"<col [^>]*min=\"28\"[^>]*outlineLevel", z.read(n).decode())]
        adj = "xl/worksheets/sheet%d.xml" % ([s.Name for s in new.Sheets].index("Adjustments") + 1)
        rules = re.findall(r'<dataValidation [^>]*sqref="(D\d+):D\d+"[^>]*>\s*<formula1>([^<]*)<',
                           z.read(adj).decode())
    c.eq("Adjustments amount validations", 4, len(rules))
    for first, f in rules:
        c.true("Adjustments %s validation reads %s" % (first, f), re.findall(r"D\d+", f) == [first, first])
    pb_index = [s.Name for s in new.Sheets].index("Projected_Balances") + 1
    c.true("Projected_Balances AB still grouped",
           "xl/worksheets/sheet%d.xml" % pb_index not in grouped)

    # Headings and labels from the user's edits.
    c.eq("Projected_Balances B13:M13", block(edits, "Projected_Balances", "B13:M13"),
         block(new, "Projected_Balances", "B13:M13"))
    c.eq("Settings A6", get(edits.Sheets.getByName("Settings"), "A6"),
         get(new.Sheets.getByName("Settings"), "A6"))
    settings = block(new, "Settings", "A10:B12")
    c.eq("Cfg_YearsBetweenRFS", 3.0, new.NamedRanges.getByName("Cfg_YearsBetweenRFS").ReferredCells.getValue())
    c.true("Settings A10 is not the new setting", settings[0][0].startswith("Years between RFS updates"))
    c.eq("Settings A11", "", settings[1][0])
    c.true("Settings B11 is unlocked",
           new.Sheets.getByName("Settings").getCellRangeByName("B11").CellProtection.IsLocked)
    c.true("Settings A12 is not the note", settings[2][0].startswith("NOTE:"))

    # README.
    readme = new.Sheets.getByName("README")
    text = [r[0] for r in block(new, "README", "A1:A40")]
    for want in (
        "https://pvillela.github.io/spatac-workbook/",
        "the Board's adjustments to the RFS",
        "and a what-if section to forecast the impact of actuals and Spatac adjustments.",
        "Settings — configuration settings, including the base year,",
        "Empty/blank means no change;",
        "the Current estimate column shows",
        "Spatac column (collapsed by default) —",
        "flags a row without S (see Spatac column above)",
        '(shows "In RFS: clear S" and the S should be deleted)',
        "The column updates by itself. After pasting a new RFS or adding a row, scan it and "
        "resolve every flag.",
    ):
        c.true("README: %r missing" % want, sum(want in t for t in text) == 1)
    c.eq("README A4", "https://pvillela.github.io/spatac-workbook/", text[3])
    for i, t in enumerate(text):
        if " — " in t:
            cur = readme.getCellByPosition(0, i).createTextCursor()
            cur.gotoStart(False)
            cur.goRight(t.index(" — "), True)
            c.true("README A%d: topic not bold" % (i + 1), cur.CharWeight == 150.0)
            if not t.startswith("Red,"):  # that line shows the Alarm style, bold throughout
                cur.gotoEnd(False)
                cur.goLeft(1, True)
                c.true("README A%d: whole line bold" % (i + 1), cur.CharWeight != 150.0)
    for prefix, color in (("Yellow", 0xFFF2CC), ("Green", 0xE2EFDA), ("Grey", 0xF2F2F2),
                          ("Orange", 0xFCE4D6), ("Light grey", 0xEDEDED), ("Red,", 0xC00000)):
        r = [i for i, t in enumerate(text) if t.startswith(prefix + " ")]
        c.true("README: one %s line" % prefix, len(r) == 1)
        if r:
            c.eq("README %s swatch" % prefix, color, readme.getCellByPosition(0, r[0]).CellBackColor)

    for tab in SPATAC_TABS:
        c.true("%s is not protected" % tab, new.Sheets.getByName(tab).isProtected())

    # Last, as it is not undone: the what-if does not read the RFS block. Clear it, and O:Z
    # (bar V, the comparison) stay.
    before = block(new, "Projected_Balances", "O14:Z63")
    pb.unprotect("")
    rfs = pb.getCellRangeByName("B14:M63")
    rfs.clearContents(1 | 2 | 4 | 16)
    new.calculateAll()
    after = block(new, "Projected_Balances", "O14:Z63")
    for i, (b0, b1) in enumerate(zip(before, after)):
        for j, (x, y) in enumerate(zip(b0, b1)):
            if col(15 + j) != "V":
                c.eq("Projected_Balances %s%d without the RFS block" % (col(15 + j), 14 + i), x, y)
    c.eq("Projected_Balances H20 cleared", "", block(new, "Projected_Balances", "H20:H20")[0][0])


def main():
    old_path, edits_path, new_path = sys.argv[1:4]
    c = Checks()
    with Session() as s:
        old, edits, new = s.open_copy(old_path), s.open_copy(edits_path), s.open_copy(new_path)
        run(c, old, edits, new, new_path)
        for d in (old, edits, new):
            d.close(True)
    print("%d checks" % c.count)
    if c.failures:
        for f in c.failures[:60]:
            print("FAIL " + f)
        print("%d failures" % len(c.failures))
        return 1
    print("all passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
