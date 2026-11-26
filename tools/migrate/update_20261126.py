"""Applies the 2026-11-26 defects and enhancements (docs/archive/design-history.md section 16).

Usage (LibreOffice's bundled Python carries the `uno` module):

    /opt/libreoffice26.2/program/python tools/migrate/update_20261126.py \
        SOURCE.xlsx OUTPUT.xlsx

SOURCE is the user's baseline, _nogit/Spatac-20261125h-PV.xlsx (Spatac-20261125h.xlsx with the
Medium_Term_Cashflow fix the DS review asked for). It is opened read-only. OUTPUT must not exist
yet: the script never overwrites a file.

Like restructure.py, the script edits SOURCE in place: what section 16 does not change keeps its
formats, comments, groups and names.
"""

import argparse
import math
import os
import re
import sys
import tempfile

import uno  # noqa: F401  (registers the com.sun.star import hook)
from com.sun.star.table import CellRangeAddress

from migrate import YEAR, Builder, define, ref
from restructure import BEYOND_COLOR, CHANGED_COLOR, RFS_SPAN, keep_autofilter, protection
from uno_session import Session

YEARS_BETWEEN_RFS = 3  # the RFS is updated every three years (design 16.6)
ALARM_COLOR = 0xC00000  # a failed check: red fill, bold white text (design 16.3)
ONE_LINE = 452  # LibreOffice's height of one line of the default font, in 1/100 mm
CHAR_WIDTH = 160  # the default font's average character width, in 1/100 mm
WIDER = 1.15  # how much wider text may render at another zoom or in Excel
FIRST_ITEM_ROW = 15  # RFS_fvalues: the item rows, below the "Item | Building Element" header
LAST_ITEM_ROW = 400

# Projected_Balances B13:M13, as the user edited them in Spatac-20261125h-PV-edits.xlsx.
RFS_HEADINGS = [
    "Opening balance", "Expected costs", "Contribution increase", "Contribution",
    "Interest rate", "Interest", "Closing balance", "Diff from RFS_cashflow balance",
    "Margin over minimum", "Margin over inflated minimum", "PV margin over minimum",
    "Contribution per unit per month",
]

PLUG_COMMENT = (
    "A plug: it rounds the first year's 5-year investment to a round amount, and brings a "
    "trivial 5-year investment (under $1,000) in a later year to zero. See the Investment "
    "Plan page of the documentation."
)


def pv(name, value):
    p = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    p.Name, p.Value = name, value
    return p


def alarm(sheet, address, formula=None, operator="FORMULA"):
    """Adds the Alarm format (red fill, bold white text) to `address`: when `formula` (relative
    to the range's first cell) is true, or, with another operator, when the value meets it."""
    rng = sheet.getCellRangeByName(address)
    cf = rng.ConditionalFormat
    cf.addNew((
        pv("Operator", uno.Enum("com.sun.star.sheet.ConditionOperator", operator)),
        pv("Formula1", formula),
        pv("StyleName", "Alarm"),
        pv("SourcePosition", rng.getCellByPosition(0, 0).CellAddress),
    ))
    rng.ConditionalFormat = cf


def slack_height(sheet, row, lines=1):
    """An explicit height: LibreOffice's fitted height, with Builder.fix_heights's margin for
    Excel's wider font metrics, plus `lines` lines of slack, so that the text still fits when it
    reflows at another zoom (design 16.2)."""
    r = sheet.Rows.getByIndex(row - 1)
    r.OptimalHeight = True
    r.Height = int(r.Height * 1.1) + lines * ONE_LINE


def bold_topic(cell):
    """Bolds the text before the first " — " (the line's topic)."""
    text = cell.getString()
    n = text.find(" — ")
    if n <= 0:
        return False
    cursor = cell.createTextCursor()
    cursor.gotoStart(False)
    cursor.goRight(n, True)
    cursor.CharWeight = 150.0
    return True


class Update:
    def __init__(self, session, src_path, out_path):
        self.s = session
        self.src_path = src_path
        self.out_path = out_path
        self.report = []

    def run(self):
        if os.path.exists(self.out_path):
            raise SystemExit("%s exists: the script never overwrites a file" % self.out_path)
        doc = self.s.open_copy(self.src_path, hidden=False)  # hidden documents lose frozen panes
        self.doc = doc
        doc.calculateAll()
        for sheet in doc.Sheets:
            if sheet.isProtected():
                sheet.unprotect("")
        self.add_styles()
        define(doc, "RFS_Fv_Items", ref("RFS_fvalues", "C%d:AZ%d" % (FIRST_ITEM_ROW, LAST_ITEM_ROW)))

        self.settings()
        self.expenditures()
        self.adjustments()
        self.mtc()
        self.projected()
        self.readme()
        self.amount_validations()
        self.protect()

        doc.calculateAll()
        with tempfile.TemporaryDirectory(prefix="spatac-update-") as tmp:
            saved = os.path.join(tmp, "out.xlsx")
            self.s.save_xlsx(doc, saved)
            keep_autofilter(saved, self.out_path)
        doc.close(True)
        return self.report

    def add_styles(self):
        styles = self.doc.StyleFamilies.getByName("CellStyles")
        if styles.hasByName("Alarm"):
            raise RuntimeError("the Alarm style exists already")
        s = self.doc.createInstance("com.sun.star.style.CellStyle")
        styles.insertByName("Alarm", s)
        s.ParentStyle = "Default"
        s.CellBackColor = ALARM_COLOR
        s.IsCellBackgroundTransparent = False
        s.CharColor = 0xFFFFFF
        s.CharWeight = 150.0

    # -- Settings -----------------------------------------------------------------------------

    def settings(self):
        """The user's label (A6); a new input, the years between RFS updates (B10); a blank row
        before the note, as in the user's edits (design 16.6)."""
        sheet = self.doc.Sheets.getByName("Settings")
        b = Builder(self.doc, sheet)
        if not b.cell("A6").getString().startswith("Year in whose dollars the minimum target is"):
            raise RuntimeError("Settings A6 is not the minimum target's year")
        b.value("A6", "Year in whose dollars the minimum target cash balance is stated")
        if not b.cell("A10").getString().startswith("NOTE:"):
            raise RuntimeError("Settings A10 is not the note")
        sheet.Rows.insertByIndex(9, 2)  # the note moves to A12
        b.value("A10", "Years between RFS updates (the what-if summary on Projected_Balances runs "
                       "this many years past the RFS)")
        b.value("B10", YEARS_BETWEEN_RFS)
        b.style("B10", "Input")
        b.fmt("B10", "0")
        b.cell("B10").HoriJustify = b.cell("B9").HoriJustify
        b.cell("A10").CellStyle = b.cell("A9").CellStyle
        b.cell("A11:B11").CellStyle = "Default"  # the inserted rows copy B9's input format
        b.cell("A11:B11").CellProtection = protection(True)
        define(self.doc, "Cfg_YearsBetweenRFS", ref("Settings", "B10"))

    # -- Expenditures -------------------------------------------------------------------------

    def expenditures(self):
        """Row 3 adds up the RFS items instead of reading the vendor's total, and is checked
        against it (design 16.4)."""
        sheet = self.doc.Sheets.getByName("Expenditures")
        b = Builder(self.doc, sheet)
        if b.cell("Q3").getString() != "RFS total":
            raise RuntimeError("Expenditures Q3 is not the RFS total label")
        if b.cell("U6").getFormula() != "=RFS_StudyYear" or b.cell("AX6").getValue() != 2053:
            raise RuntimeError("Expenditures U6:AX6 is not the RFS's 30 years")
        for a in ("L3", "M3", "N3", "O3", "P3"):
            if b.cell(a).getString() or b.cell(a).getFormula():
                raise RuntimeError("Expenditures %s is not empty" % a)
        fv = self.doc.Sheets.getByName("RFS_fvalues")
        if fv.getCellRangeByName("A%d" % (FIRST_ITEM_ROW - 1)).getString() != "Item":
            raise RuntimeError("RFS_fvalues A%d is not the Item header" % (FIRST_ITEM_ROW - 1))

        b.value("Q3", "RFS total (sum of items)")
        b.formulas("U3:AX3", [[
            "=SUM(INDEX(RFS_Fv_Items;0;MATCH(%s$6;RFS_Fv_Years;0)))" % c
            for c in self.columns("U", "AX")
        ]])
        alarm(sheet, "U3:AX3", "ABS(U3+INDEX(RFS_Fv_Totals;1;MATCH(U$6;RFS_Fv_Years;0)))>0.5")
        right = uno.Enum("com.sun.star.table.CellHoriJustify", "RIGHT")
        b.value("O3", "Years where items differ from RFS_fvalues totals")
        b.cell("O3").HoriJustify = right
        b.formula("P3", "=SUMPRODUCT(--(ABS(U3:AX3+SUMIF(RFS_Fv_Years;U6:AX6;RFS_Fv_Totals))>0.5))")
        b.fmt("P3", "0")
        alarm(sheet, "P3", "0", "GREATER")

    @staticmethod
    def columns(first, last):
        from migrate import col, colnum

        return [col(k) for k in range(colnum(first), colnum(last) + 1)]

    # -- Adjustments --------------------------------------------------------------------------

    def adjustments(self):
        """Header rows tall enough at any zoom (design 16.2); the to-year amount no longer hides
        an error (design 16.5)."""
        sheet = self.doc.Sheets.getByName("Adjustments")
        b = Builder(self.doc, sheet)
        data = b.cell("A1:A200").getDataArray()
        headers = [i + 1 for i, (v,) in enumerate(data) if v == "Item (GL account - element)"]
        if len(headers) != 4:
            raise RuntimeError("Adjustments: %d block headers, not 4" % len(headers))
        for r in headers:
            slack_height(sheet, r)
            if b.cell("A%d" % (r + 11)).getString() or not b.cell("D%d" % (r + 1)).Validation.Formula1:
                raise RuntimeError("Adjustments: the block below row %d is not 10 input rows" % r)
        self.adj_headers = headers
        changed = 0
        for r in range(1, 201):
            c = b.cell("G%d" % r)
            f = c.getFormula()
            if not f.startswith("=IFERROR("):
                continue
            m = re.fullmatch(r'=IFERROR\((LET\(.*\));\s*""\)', f, re.S)
            if not m:
                raise RuntimeError("Adjustments G%d: unexpected formula %r" % (r, f))
            c.setFormula("=" + m.group(1))
            changed += 1
        if changed != 40:
            raise RuntimeError("Adjustments: %d to-year formulas, not 40" % changed)

    def amount_validations(self):
        """Inserting rows on Settings moves the relative reference of the Adjustments amount
        validations below the inserted rows (a LibreOffice bug): D23:D32 came to check D25.
        Each block's validation is set again, last, from its first cell."""
        sheet = self.doc.Sheets.getByName("Adjustments")
        for r in self.adj_headers:
            cells = sheet.getCellRangeByName("D%d:D%d" % (r + 1, r + 10))
            v = cells.Validation
            v.setSourcePosition(cells.getCellByPosition(0, 0).CellAddress)
            v.Formula1 = "AND(ISNUMBER(D%d);D%d>0)" % (r + 1, r + 1)
            cells.Validation = v

    # -- Medium_Term_Cashflow -----------------------------------------------------------------

    def mtc(self):
        sheet = self.doc.Sheets.getByName("Medium_Term_Cashflow")
        b = Builder(self.doc, sheet)
        if b.cell("A14").getString() != "1-year maturity" or \
                b.cell("A18").getString() != "5-year maturity (balancing figure)":
            raise RuntimeError("Medium_Term_Cashflow A14:A18 is not the investment grid")
        if b.cell("A23").getString() != "Adjustment to minimum cash balance":
            raise RuntimeError("Medium_Term_Cashflow A23 is not the adjustment row")
        alarm(sheet, "B14:G18", "0", "LESS")  # a negative investment (design 16.7)
        sheet.Annotations.insertNew(b.cell("A23").CellAddress, PLUG_COMMENT)

    # -- Projected_Balances -------------------------------------------------------------------

    def projected(self):
        doc = self.doc
        sheet = doc.Sheets.getByName("Projected_Balances")
        b = Builder(doc, sheet)
        hdr, first, last = 13, 14, 63
        last30 = first + RFS_SPAN - 1
        if b.cell("I%d" % hdr).getString() != "Difference from RFS_cashflow balance":
            raise RuntimeError("Projected_Balances I13 is not the difference column")
        if b.cell("H12").getString() != "Years that differ":
            raise RuntimeError("Projected_Balances H12 is not the check's label")

        b.value("A2", "The RFS's projection recomputed from its inputs (columns B-M), checked "
                      "against RFS_cashflow (the costs in column C, the balances in column I), "
                      "and a what-if projection (columns O-Z, inputs in P4:P10).")
        b.values("B%d:M%d" % (hdr, hdr), [RFS_HEADINGS])
        b.value("B12", "RFS (recomputed)")

        rows = range(first, last + 1)
        # The RFS block (design 16.4, 16.5): the costs add up the items of RFS_fvalues; neither
        # lookup into RFS_cashflow hides an error, since it covers all 50 years.
        b.formulas("C%d:C%d" % (first, last), [
            ["=-SUM(INDEX(RFS_Fv_Items;0;MATCH($A%d;RFS_Fv_Years;0)))" % r] for r in rows])
        b.formulas("I%d:I%d" % (first, last), [
            ["=$H{r}-INDEX(RFS_Cf_Balance;MATCH($A{r};RFS_Cf_Year;0))".format(r=r)] for r in rows])

        # The what-if reads the RFS's own figures, never the RFS block (design 16.5).
        p_col, q_col, r_col, s_col = [], [], [], []
        for r in rows:
            p = r - 1
            p_col.append([
                "=LET(\n"
                "  year; $A{r};\n"
                "  IFS(\n"
                '    $O{r}=""; "";\n'
                "    year=Cfg_BaseYear; -SUM(Exp_Current);\n"
                "    year>MAX(Exp_AdjustedYears); -SUM(INDEX(RFS_Fv_Items; 0; MATCH(year; RFS_Fv_Years; 0)));\n"
                "    TRUE(); -SUM(INDEX(Exp_Adjusted; 0; MATCH(year; Exp_AdjustedYears; 0)))))".format(r=r)
            ])
            q_col.append([
                "=LET(\n"
                "  year; $A{r};\n"
                "  n; year-RFS_StudyYear;\n"
                "  years_after; year-PB_ContribFixedThrough;\n"
                "  IFS(\n"
                '    OR($O{r}=""; n=0); "";\n'
                "    years_after>3; PB_SubsequentIncrease;\n"
                "    years_after>0; INDEX(PB_CatchUp; years_after);\n"
                "    n<=3; INDEX(RFS_CatchUpIncreases; n);\n"
                "    TRUE(); RFS_SubsequentIncrease))".format(r=r)
            ])
            r_col.append([
                '=IF($O{r}=""; ""; RFS_PresentContribution)'.format(r=r) if r == first else
                "=LET(\n"
                "  year; $A{r};\n"
                "  IFS(\n"
                '    $O{r}=""; "";\n'
                "    year=RFS_StudyYear; RFS_PresentContribution;\n"
                "    year=PB_StartYear; INDEX(RFS_Cf_Contrib; MATCH(year-1; RFS_Cf_Year; 0))*(1+$Q{r});\n"
                "    TRUE(); $R{p}*(1+$Q{r})))".format(r=r, p=p)
            ])
            s_col.append([
                "=IFS(\n"
                '  $O{r}=""; "";\n'
                "  $A{r}<RFS_StudyYear+Cfg_ShortInterestYears; RFS_ShortInterest;\n"
                "  TRUE(); RFS_LongInterest)".format(r=r)
            ])
        b.formulas("P%d:P%d" % (first, last), p_col)
        b.formulas("Q%d:Q%d" % (first, last), q_col)
        b.formulas("R%d:R%d" % (first, last), r_col)
        b.formulas("S%d:S%d" % (first, last), s_col)

        # The checks (design 16.3, 16.4): the costs against RFS_cashflow, beside the balances'.
        right = uno.Enum("com.sun.star.table.CellHoriJustify", "RIGHT")
        b.value("E12", "Cost years that differ")
        b.cell("E12").HoriJustify = right
        b.formula("F12", "=SUMPRODUCT(--(ABS(C%d:C%d-SUMIF(RFS_Cf_Year;A%d:A%d;RFS_Cf_Cost))>0.5))"
                  % (first, last, first, last))
        b.value("H12", "Balance years that differ")
        b.bold("E12:I12")
        for a in ("F12", "I12"):
            b.fmt(a, "0")
            alarm(sheet, a, "0", "GREATER")
        alarm(sheet, "C%d:C%d" % (first, last),
              "ABS($C%d-INDEX(RFS_Cf_Cost;MATCH($A%d;RFS_Cf_Year;0)))>0.5" % (first, first))
        # The per-year balance flag becomes as loud as the counts.
        cells = sheet.getCellRangeByName("I%d:I%d" % (first, last))
        cf = cells.ConditionalFormat
        cf.clear()
        cells.ConditionalFormat = cf
        alarm(sheet, "I%d:I%d" % (first, last), "ABS(N(I%d))>0.5" % first)

        # The what-if summary: from the current year through the RFS's last year plus the years
        # to the next RFS (design 16.6).
        self.what_if_summary(b, first, last)

        # The RFS block is not green, to set it apart from the what-if (design 16.8); the years
        # past the RFS keep their light grey.
        fmts = [[sheet.getCellByPosition(c, r - 1).NumberFormat for c in range(1, 13)] for r in rows]
        b.style("B%d:M%d" % (first, last), "Default")
        b.cell("B%d:M%d" % (first, last)).IsCellBackgroundTransparent = True
        for i, r in enumerate(rows):
            for c in range(1, 13):
                sheet.getCellByPosition(c, r - 1).NumberFormat = fmts[i][c - 1]
        b.fill("A%d:Z%d" % (last30 + 1, last), BEYOND_COLOR)
        # The first year's row had kept the header's height since blank row 11 went (design
        # 15.7).
        sheet.Rows.getByIndex(first - 1).OptimalHeight = True

        # The inflator always shows (design 16.8).
        a = CellRangeAddress()
        a.Sheet = sheet.RangeAddress.Sheet
        a.StartColumn = a.EndColumn = 27  # AB
        a.StartRow = a.EndRow = 0
        sheet.showDetail(a)
        sheet.ungroup(a, uno.Enum("com.sun.star.table.TableOrientation", "COLUMNS"))
        sheet.Columns.getByName("AB").IsVisible = True

    def what_if_summary(self, b, first, last):
        years = "$A$%d:$A$%d" % (first, last)
        span = (
            "  first; Cfg_BaseYear;\n"
            "  last; RFS_StudyYear+%d+Cfg_YearsBetweenRFS;\n" % (RFS_SPAN - 1)
        )
        b.formula("R4", '="What-if, "&Cfg_BaseYear&"–"&(RFS_StudyYear+%d+Cfg_YearsBetweenRFS)'
                  % (RFS_SPAN - 1))
        for r, m in ((5, "W"), (6, "X"), (7, "Y")):
            margins = "%s$%d:%s$%d" % (m, first, m, last)
            within = '{y}; ">="&first; {y}; "<="&last'.format(y=years)
            b.formula("T%d" % r, "=LET(\n" + span + "  MINIFS(%s; %s))" % (margins, within))
            b.formula("U%d" % r, "=LET(\n" + span + "  MINIFS(%s; %s; T%d; %s))"
                      % (years, margins, r, within))
            b.formula(
                "W%d" % r,
                "=LET(\n" + span +
                "  negatives; COUNTIFS(%s; \"<0\"; %s);\n" % (margins, within) +
                '  IF(negatives=0; "none"; MINIFS(%s; %s; "<0"; %s)))' % (years, margins, within),
            )
            b.formula("V%d" % r, '=IF(W{r}="none"; ""; INDEX({m}; MATCH(W{r}; {y}; 0)))'
                      .format(r=r, m=margins, y=years))
            b.fmt("U%d" % r, YEAR)
            b.fmt("W%d" % r, YEAR)

    # -- README -------------------------------------------------------------------------------

    def readme(self):
        sheet = self.doc.Sheets.getByName("README")
        b = Builder(self.doc, sheet)
        a3 = b.cell("A3").getString()
        url = "https://pvillela.github.io/spatac-workbook/"
        if not a3.startswith("Documentation: " + url):
            raise RuntimeError("README A3 is not the documentation line")
        # The URL gets a row of its own (design 16.2).
        sheet.Rows.insertByIndex(3, 1)
        b.value("A3", "Documentation — what each tab is for, how to use it, and the January and "
                      "new-RFS procedures:")
        b.value("A4", url)
        b.cell("A4").CellStyle = b.cell("A3").CellStyle
        b.cell("A4").IsTextWrapped = True

        def line(prefix):
            data = b.cell("A1:A40").getDataArray()
            rows = [i + 1 for i, (v,) in enumerate(data) if str(v).startswith(prefix)]
            if len(rows) != 1:
                raise RuntimeError("README: %r found %d times" % (prefix, len(rows)))
            return rows[0]

        def edit(prefix, old, new):
            r = line(prefix)
            text = b.cell("A%d" % r).getString()
            if text.count(old) != 1:
                raise RuntimeError("README A%d: %r not found once" % (r, old))
            b.cell("A%d" % r).setString(text.replace(old, new))
            return r

        edit("Adjustments — ", "the Board's changes to the RFS", "the Board's adjustments to the RFS")
        edit("Projected_Balances — ", "and a what-if with Spatac's adjustments and contribution "
             "increases.", "and a what-if section to forecast the impact of actuals and Spatac "
             "adjustments.")
        edit("Settings — ", "Settings — the base year,", "Settings — configuration settings, "
             "including the base year,")
        edit("Settings — ", "of short-term inflation and of short-term interest.",
             "of short-term inflation and of short-term interest, and the years between RFS "
             "updates.")
        edit("Mid-year and Late-year", "Blank means no change;", "Empty/blank means no change;")
        edit("Mid-year and Late-year", "the Current estimate columns show",
             "the Current estimate column shows")
        edit("Spatac row — ", "Spatac row —", "Spatac column (collapsed by default) —")
        edit("RFS check — ", "flags a row without S", "flags a row without S (see Spatac column above)")
        edit("RFS check — ", '("In RFS: clear S")',
             '(shows "In RFS: clear S" and the S should be deleted)')
        edit("RFS check — ", "Check it after pasting a new RFS or adding a row.",
             "The column updates by itself. After pasting a new RFS or adding a row, scan it and "
             "resolve every flag.")

        # Each colour's line shows the colour (design 16.2), and a line for the new one.
        orange, grey = line("Orange — "), line("Light grey — ")
        for r, color in ((orange, CHANGED_COLOR), (grey, BEYOND_COLOR)):
            b.cell("A%d" % r).CellStyle = "Default"
            b.fill("A%d" % r, color)
            b.wrap("A%d" % r)
        sheet.Rows.insertByIndex(grey, 1)
        red = grey + 1
        b.value("A%d" % red, "Red, bold white text — a check that found a discrepancy: a count that "
                             "is not 0, or the value it counts.")
        b.cell("A%d" % red).CellStyle = "Alarm"
        b.wrap("A%d" % red)

        # Bold topics, and explicit heights with a line of slack for every line that wraps.
        # A line of slack only where the text ends near a line's end: those are the lines that
        # reflow onto one more line at another zoom (A3's did, at 110%).
        used = b.cell("A1:A40").getDataArray()
        per_line = sheet.Columns.getByIndex(0).Width / CHAR_WIDTH
        bolded, slack = 0, 0
        for i, (text,) in enumerate(used):
            r = i + 1
            if not text or r == 1:
                continue
            if bold_topic(b.cell("A%d" % r)):
                bolded += 1
            near_end = math.ceil(len(text) * WIDER / per_line) > math.ceil(len(text) / per_line)
            slack_height(sheet, r, 1 if near_end else 0)
            slack += near_end
        self.report.append("README: %d topics in bold, %d lines with a line of slack" % (bolded, slack))

    # -- protection ---------------------------------------------------------------------------

    def protect(self):
        sheets = self.doc.Sheets
        for name in ("README", "Settings", "Expenditures", "Adjustments", "Medium_Term_Cashflow",
                     "Schedule_of_Investments", "Projected_Balances"):
            sheets.getByName(name).protect("")
        self.doc.CurrentController.setActiveSheet(sheets.getByName("README"))


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("source")
    ap.add_argument("output")
    args = ap.parse_args()
    with Session() as s:
        report = Update(s, args.source, args.output).run()
    print("Wrote %s" % args.output)
    for line in report:
        print("- " + line)


if __name__ == "__main__":
    sys.exit(main())
