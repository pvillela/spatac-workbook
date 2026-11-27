"""Checks update_20261127.py's output (design-history.md section 17).

Usage:

    /opt/libreoffice26.2/program/python tools/migrate/verify_20261127.py \
        BUILD.xlsx EDITS.xlsx NEW.xlsx

BUILD is the previous build (_nogit/Spatac-20261126a.xlsx); EDITS is the user's edited copy of
it and the baseline (_nogit/Spatac-20261126a-edits.xlsx); NEW is the output. All are opened
read-only. The checks:

- the user's edits changed no value of BUILD, allowing for the column the user inserted on
  Expenditures, apart from the cells they edited;
- every value of NEW equals EDITS's, apart from the cells section 17 changes on purpose;
- no error values on any Spatac tab;
- the layout check (layout_lint.py) finds nothing;
- the moved check (Expenditures T1) is 0 today, catches an item that no longer adds up, and has
  the Alarm format; P3 has none;
- the alarms on Projected_Balances C and I show white text on red (the saved file's formats);
- the new headings, wrap and row heights, protection of every tab;
- the regressions: a check for each defect of the earlier defect files that a build can bring
  back.
"""

import re
import sys
import zipfile

import uno  # noqa: F401
from layout_lint import Lint
from uno_session import Session
from verify_20261126 import alarm_at, compare, no_errors
from verify_restructure import Checks, block, col, colnum

SPATAC_TABS = ["README", "Expenditures", "Adjustments", "Medium_Term_Cashflow",
               "Schedule_of_Investments", "Projected_Balances", "Settings"]
VENDOR_TABS = ["RFS_info", "RFS_fvalues", "RFS_cashflow"]
INSERTED = colnum("T")  # the column the user inserted on Expenditures
HEADINGS = {"L": "Approved", "M": "Actuals to date", "N": "Mid-year estimate",
            "O": "Late-year estimate", "P": "Current estimate"}


def get(doc, sheet, a):
    return block(doc, sheet, "%s:%s" % (a, a))[0][0]


def user_edits(c, build, edits):
    """The user's edits (design 17.1) moved Expenditures T: onwards right by one column and
    changed only labels."""
    def shifted(a):
        m = re.fullmatch(r"([A-Z]+)(\d+)", a)
        n = colnum(m.group(1))
        return "%s%s" % (col(n + 1) if n >= INSERTED else m.group(1), m.group(2))

    o, n = build.Sheets.getByName("Expenditures"), edits.Sheets.getByName("Expenditures")
    od = o.getCellRangeByPosition(0, 0, colnum("AX") - 1, 249).getDataArray()
    nd = n.getCellRangeByPosition(0, 0, colnum("AY") - 1, 249).getDataArray()
    skip = {"O3", "P3", "T1"}  # the count left O3:P3; the user labelled it in U1 (T1 before)
    for r, row in enumerate(od):
        for k, v in enumerate(row):
            a = "%s%d" % (col(k + 1), r + 1)
            if a in skip:
                continue
            b = shifted(a)
            m = re.fullmatch(r"([A-Z]+)(\d+)", b)
            c.eq("edits: Expenditures %s -> %s" % (a, b), v, nd[int(m.group(2)) - 1][colnum(m.group(1)) - 1])
    c.eq("edits: Expenditures T1 (the moved count)", get(build, "Expenditures", "P3"),
         get(edits, "Expenditures", "T1"))
    for tab in ("Adjustments", "Medium_Term_Cashflow", "Schedule_of_Investments",
                "Projected_Balances", "Settings") + tuple(VENDOR_TABS):
        compare(c, build, edits, tab, skip={"D6", "D22", "D38", "D54"} if tab == "Adjustments" else ())
    compare(c, build, edits, "README", skip={"A21"})


def run(c, build, edits, new, new_path, lint_doc):
    for d in (build, edits, new):
        d.calculateAll()

    # Values.
    user_edits(c, build, edits)
    changed = {"Expenditures": {"%s7" % k for k in HEADINGS}}
    for tab in SPATAC_TABS + VENDOR_TABS:
        compare(c, edits, new, tab, skip=changed.get(tab, ()))
    no_errors(c, new)

    # The layout check.
    problems = Lint(lint_doc).run()
    for p in problems:
        c.true("layout: " + p, False)
    c.true("layout check ran", True)

    exp = new.Sheets.getByName("Expenditures")
    pb = new.Sheets.getByName("Projected_Balances")
    fv = new.Sheets.getByName("RFS_fvalues")

    # The moved check.
    c.eq("Expenditures T1 today", 0.0, get(new, "Expenditures", "T1"))
    c.true("Expenditures T1 has no Alarm format", alarm_at(new, exp, "T1"))
    c.eq("Expenditures P3 conditional formats", 0, exp.getCellRangeByName("P3").ConditionalFormat.Count)
    c.eq("Expenditures U1", "  No. of RFS years where totals don’t add-up correctly",
         get(new, "Expenditures", "U1"))
    cell = fv.getCellRangeByName("C22")
    kept = cell.getValue()
    cell.setValue(kept + 1000)
    new.calculateAll()
    c.eq("after RFS_fvalues C22 + 1000: Expenditures T1", 1.0, get(new, "Expenditures", "T1"))
    cell.setValue(kept)
    new.calculateAll()

    # Medium_Term_Cashflow row 11: no IFERROR, same values (the value comparison covers them).
    mtc = new.Sheets.getByName("Medium_Term_Cashflow")
    for k in "BCDEFG":
        f = mtc.getCellRangeByName("%s11" % k).getFormula()
        c.true("Medium_Term_Cashflow %s11 hides errors: %s" % (k, f[:60]), "IFERROR" not in f)
        c.true("Medium_Term_Cashflow %s11 is not the LET form" % k, "SUMPRODUCT(amounts;" in f)
        c.true("Medium_Term_Cashflow %s11 holds an array in a LET name" % k, "term;" not in f)
    c.eq("Medium_Term_Cashflow B11:G11",
         [0.0, 0.0, 399619.4, 0.0, 103000.0, 646492.67],
         [round(v, 2) for v in block(new, "Medium_Term_Cashflow", "B11:G11")[0]])

    # Row 7: the year on a line of its own.
    for k, label in HEADINGS.items():
        c.eq("Expenditures %s7" % k, "2026\n" + label, get(new, "Expenditures", "%s7" % k))

    # Projected_Balances C and I: the alarm's own number format, without a colour code, so that
    # its white text shows (read from the saved file: that is what Excel and Calc will show).
    with zipfile.ZipFile(new_path) as z:
        names = [s.Name for s in new.Sheets]
        sheet_xml = z.read("xl/worksheets/sheet%d.xml" % (names.index("Projected_Balances") + 1)).decode()
        styles = z.read("xl/styles.xml").decode()
    dxfs = re.findall(r"<dxf>(.*?)</dxf>", re.search(r"<dxfs[^>]*>(.*?)</dxfs>", styles, re.S).group(1), re.S)
    for rng in ("C14:C63", "I14:I63"):
        m = re.search(r'<conditionalFormatting sqref="%s"><cfRule [^>]*dxfId="(\d+)"' % rng, sheet_xml)
        c.true("Projected_Balances %s: no alarm" % rng, m is not None)
        if m:
            dxf = dxfs[int(m.group(1))]
            c.true("Projected_Balances %s: alarm not white bold on red" % rng,
                   'rgb="FFFFFFFF"' in dxf and "<b val=\"1\"" in dxf and "FFC00000" in dxf)
            fmt = re.search(r'<numFmt [^>]*formatCode="([^"]*)"', dxf)
            c.true("Projected_Balances %s: alarm without its own number format" % rng, fmt is not None)
            if fmt:
                c.true("Projected_Balances %s: alarm number format %r has a colour" % (rng, fmt.group(1)),
                       "[" not in fmt.group(1))
    c.true("Projected_Balances C14 has no Alarm format", alarm_at(new, pb, "C14"))
    c.eq("Projected_Balances I12", "=SUMPRODUCT(--(I14:I63<>0))", pb.getCellRangeByName("I12").getFormula())
    c.eq("Projected_Balances I12 today", 0.0, get(new, "Projected_Balances", "I12"))
    bal = new.Sheets.getByName("RFS_cashflow").getCellRangeByName("G20")
    kept = bal.getValue()
    for delta, want in ((1000, 1.0), (0.4, 0.0), (-0.6, 1.0)):
        bal.setValue(kept + delta)
        new.calculateAll()
        c.eq("after RFS_cashflow G20 %+g: Projected_Balances I12" % delta, want,
             get(new, "Projected_Balances", "I12"))
    bal.setValue(kept)
    new.calculateAll()
    for r in (14, 33, 63):
        c.true("Projected_Balances I%d is not rounded" % r,
               pb.getCellRangeByName("I%d" % r).getFormula().startswith("=ROUND("))

    # Wrap and heights.
    c.true("Expenditures C12 does not wrap", exp.getCellRangeByName("C12").IsTextWrapped)
    c.true("Expenditures C250 (the marker) wraps", not exp.getCellRangeByName("C250").IsTextWrapped)
    for k in "EFGHJ":
        c.true("Expenditures %s100 wraps" % k, not exp.getCellRangeByName("%s100" % k).IsTextWrapped)
    for name, column, row in (("RFS_info", "C", 15), ("RFS_info", "D", 15), ("RFS_fvalues", "B", 30)):
        sheet = new.Sheets.getByName(name)
        c.true("%s %s%d does not wrap" % (name, column, row),
               sheet.getCellRangeByName("%s%d" % (column, row)).IsTextWrapped)
        c.true("%s row %d has a fixed height" % (name, row), sheet.Rows.getByIndex(row - 1).OptimalHeight)
    c.true("README A16 is unlocked",
           new.Sheets.getByName("README").getCellRangeByName("A16").CellProtection.IsLocked)

    for tab in SPATAC_TABS + VENDOR_TABS:
        c.true("%s is not protected" % tab, new.Sheets.getByName(tab).isProtected())

    regressions(c, new)


def regressions(c, new):
    """One check for each defect of the earlier defect files that a build can bring back."""
    def text_anywhere(pattern):
        found = []
        for tab in SPATAC_TABS:
            sheet = new.Sheets.getByName(tab)
            cur = sheet.createCursor()
            cur.gotoEndOfUsedArea(False)
            a = cur.RangeAddress
            for r, row in enumerate(sheet.getCellRangeByPosition(0, 0, a.EndColumn, a.EndRow).getDataArray()):
                for k, v in enumerate(row):
                    if isinstance(v, str) and re.search(pattern, v):
                        found.append("%s!%s%d" % (tab, col(k + 1), r + 1))
        return found

    exp = new.Sheets.getByName("Expenditures")
    # 2026-11-21: top alignment everywhere.
    for tab in SPATAC_TABS:
        sheet = new.Sheets.getByName(tab)
        cur = sheet.createCursor()
        cur.gotoEndOfUsedArea(False)
        a = cur.RangeAddress
        bad = [rng.AbsoluteName for rng in
               sheet.getCellRangeByPosition(0, 0, a.EndColumn, a.EndRow).getCellFormatRanges()
               if rng.VertJustify != 1]
        c.true("21-11 %s: not top-aligned: %s" % (tab, bad[:3]), not bad)
    # 2026-11-21: item, element, GL and its name are not editable.
    for k in "BCDE":
        c.true("21-11 Expenditures %s20 is unlocked" % k,
               exp.getCellRangeByName("%s20" % k).CellProtection.IsLocked)
    c.true("21-11 'latest estimate' is back: %s" % text_anywhere(r"(?i)latest estimate"),
           not text_anywhere(r"(?i)latest estimate"))
    c.eq("21-11 Settings is the last tab", "Settings", [s.Name for s in new.Sheets][-1])
    # 2026-11-24.
    c.true("24-11 'Vendor total' is back", not text_anywhere(r"(?i)vendor total"))
    s_vals = [v for (v,) in block(new, "Expenditures", "S8:S249")]
    c.true("24-11 Expenditures S full of zeros", sum(1 for v in s_vals if v == 0.0 and v != "") <= 3)
    for a in ("Q4", "Q5"):
        c.eq("24-11 Expenditures %s alignment" % a, "RIGHT", exp.getCellRangeByName(a).HoriJustify.value)
    c.eq("24-11 Expenditures S2", "2027", exp.getCellRangeByName("S2").getString())
    c.eq("24-11 Expenditures S2 bold", 150.0, exp.getCellRangeByName("S2").CharWeight)
    c.eq("24-11 Adjustments H39 alignment", "RIGHT",
         new.Sheets.getByName("Adjustments").getCellRangeByName("H39").HoriJustify.value)
    # 2026-11-25: red negatives on Projected_Balances; no leftover labels.
    pb = new.Sheets.getByName("Projected_Balances")
    for k in "BCEGHIJKLOPRTUVWXY":
        code = new.NumberFormats.getByKey(pb.getCellRangeByName("%s20" % k).NumberFormat).FormatString
        c.true("25-11 Projected_Balances %s20: negatives not red (%s)" % (k, code), "[RED]" in code.upper())
    c.true("25-11 leftover labels", not text_anywhere(r"Unlabelled value|Historical inflation"))
    # 2026-11-26: the inflator stays visible; the RFS block is not green.
    c.true("26-11 Projected_Balances AB hidden", pb.Columns.getByName("AB").IsVisible)
    c.true("26-11 Projected_Balances B20 green", pb.getCellRangeByName("B20").CellBackColor != 0xE2EFDA)


def main():
    build_path, edits_path, new_path = sys.argv[1:4]
    c = Checks()
    with Session() as s:
        build, edits, new = s.open_copy(build_path), s.open_copy(edits_path), s.open_copy(new_path)
        lint_doc = s.open_copy(new_path)
        lint_doc.calculateAll()
        run(c, build, edits, new, new_path, lint_doc)
        for d in (build, edits, new, lint_doc):
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
