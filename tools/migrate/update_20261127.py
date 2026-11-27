"""Applies the 2026-11-27 defects (docs/archive/design-history.md section 17).

Usage (LibreOffice's bundled Python carries the `uno` module):

    /opt/libreoffice26.2/program/python tools/migrate/update_20261127.py \
        SOURCE.xlsx OUTPUT.xlsx

SOURCE is the user's baseline, _nogit/Spatac-20261126a-edits.xlsx: Spatac-20261126a.xlsx with
the user's edits (design 17.1). It is opened read-only. OUTPUT must not exist yet: the script
never overwrites a file.

The script edits SOURCE in place: what section 17 does not change keeps its formats, comments,
groups and names. Column widths and row heights come from layout_lint.py's measurements, so
that the build passes the same layout check the verifier runs.
"""

import argparse
import os
import sys
import tempfile

import uno  # noqa: F401  (registers the com.sun.star import hook)

from layout_lint import ALARM_FILL, VENDOR_TABS, Lint
from restructure import keep_autofilter, protection
from uno_session import Session

PLAIN_NUMBER = "#,##0;-#,##0"  # the alarm's number format: no colour code to override it
ALARM_NUMBER = "AlarmNumber"
SPATAC_TABS = ("README", "Expenditures", "Adjustments", "Medium_Term_Cashflow",
               "Schedule_of_Investments", "Projected_Balances", "Settings")
FIRST_ROW, LAST_ROW = 8, 250  # Expenditures' item rows


def pv(name, value):
    p = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    p.Name, p.Value = name, value
    return p


def add_condition(rng, style, formula, operator):
    cf = rng.ConditionalFormat
    cf.addNew((
        pv("Operator", uno.Enum("com.sun.star.sheet.ConditionOperator", operator)),
        pv("Formula1", formula),
        pv("StyleName", style),
        pv("SourcePosition", rng.getCellByPosition(0, 0).CellAddress),
    ))
    rng.ConditionalFormat = cf


def clear_conditions(rng):
    cf = rng.ConditionalFormat
    cf.clear()
    rng.ConditionalFormat = cf


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
        self.lint = Lint(doc)
        self.lint.open_scratch()
        self.styles()

        self.readme()
        self.expenditures()
        self.projected()
        self.mtc()
        self.schedule()
        self.settings()
        self.vendor_tabs()
        self.input_rows()

        self.lint.close_scratch()
        self.protect()
        doc.calculateAll()
        with tempfile.TemporaryDirectory(prefix="spatac-update-") as tmp:
            saved = os.path.join(tmp, "out.xlsx")
            self.s.save_xlsx(doc, saved)
            keep_autofilter(saved, self.out_path)
        doc.close(True)
        return self.report

    def styles(self):
        """The Alarm style, as the .xlsx round trip renamed it, and a copy of it with a number
        format of its own (design 17.4)."""
        styles = self.doc.StyleFamilies.getByName("CellStyles")
        alarms = [n for n in styles.ElementNames
                  if styles.getByName(n).CellBackColor == ALARM_FILL
                  and styles.getByName(n).CharColor == 0xFFFFFF]
        if not alarms:
            raise RuntimeError("no Alarm style")
        self.alarm = alarms[0]
        if styles.hasByName(ALARM_NUMBER):
            raise RuntimeError("the %s style exists already" % ALARM_NUMBER)
        s = self.doc.createInstance("com.sun.star.style.CellStyle")
        styles.insertByName(ALARM_NUMBER, s)
        s.ParentStyle = "Default"
        s.CellBackColor = ALARM_FILL
        s.IsCellBackgroundTransparent = False
        s.CharColor = 0xFFFFFF
        s.CharWeight = 150.0
        formats = self.doc.NumberFormats
        locale = uno.createUnoStruct("com.sun.star.lang.Locale")
        key = formats.queryKey(PLAIN_NUMBER, locale, False)
        if key == -1:
            key = formats.addNew(PLAIN_NUMBER, locale)
        s.NumberFormat = key

    # -- layout helpers -----------------------------------------------------------------------

    def fit_columns(self, sheet, ranges, uniform=False):
        """Widens each column of `ranges` to what its cells need (never narrows). With
        `uniform`, all the columns take the widest need."""
        needs = {}
        for address in ranges:
            rng = sheet.getCellRangeByName(address)
            a = rng.RangeAddress
            for c in range(a.StartColumn, a.EndColumn + 1):
                for r in range(a.StartRow, a.EndRow + 1):
                    cell = sheet.getCellByPosition(c, r)
                    if cell.getString().strip() == "" or cell.getError():
                        continue
                    needs[c] = max(needs.get(c, 0), self.lint.needed_width(cell))
        if uniform and needs:
            top = max(needs.values())
            needs = {c: top for c in needs}
        widened = []
        for c, need in sorted(needs.items()):
            column = sheet.Columns.getByIndex(c)
            if column.Width < need:
                column.Width = need
                widened.append(column.Name)
        if widened:
            self.report.append("%s: widened %s" % (sheet.Name, ", ".join(widened)))

    def fit_span(self, sheet, address, first, rest):
        """Widens column `first` so that the cell's text fits `first` and the columns `rest`."""
        need = self.lint.needed_width(sheet.getCellRangeByName(address))
        others = sum(sheet.Columns.getByName(c).Width for c in rest)
        column = sheet.Columns.getByName(first)
        if column.Width + others < need:
            column.Width = need - others
            self.report.append("%s: widened %s for %s" % (sheet.Name, first, address))

    def fit_fixed_row(self, sheet, row):
        """A fixed-height row (a heading) as tall as its wrapped cells need with the margin."""
        r = sheet.Rows.getByIndex(row - 1)
        need = r.Height
        cols = sheet.Columns
        for c in range(0, 60):
            cell = sheet.getCellByPosition(c, row - 1)
            if cell.IsTextWrapped and cell.getString() and cols.getByIndex(c).IsVisible:
                need = max(need, self.lint.needed_height(cell, cols.getByIndex(c).Width))
        if need > r.Height:
            r.Height = need
            self.report.append("%s: row %d made taller" % (sheet.Name, row))

    # -- README -------------------------------------------------------------------------------

    def readme(self):
        """The Yellow legend line is a sample of the input colour, not an input (design 17.3)."""
        sheet = self.doc.Sheets.getByName("README")
        cells = sheet.getCellRangeByName("A1:A60")
        cells.CellProtection = protection(True)

    # -- Expenditures -------------------------------------------------------------------------

    def expenditures(self):
        sheet = self.doc.Sheets.getByName("Expenditures")
        cell = sheet.getCellRangeByName
        if not cell("T1").getFormula().startswith("=SUMPRODUCT(--(ABS(V3:AY3"):
            raise RuntimeError("Expenditures T1 is not the RFS totals check")
        if cell("P3").getFormula() or cell("O3").getString():
            raise RuntimeError("Expenditures O3:P3 is not empty")
        if cell("V6").getFormula() != "=RFS_StudyYear":
            raise RuntimeError("Expenditures V6 is not the RFS's first year")

        # The check moved to T1 with the user's edit; its alarm follows it (design 17.4).
        clear_conditions(cell("P3"))
        add_condition(cell("T1"), self.alarm, "0", "GREATER")

        # Row 7: the year on a line of its own, so that "Mid-year" and "Late-year" never break
        # at their hyphen (design 17.5).
        labels = {"L": "Approved", "M": "Actuals to date", "N": "Mid-year estimate",
                  "O": "Late-year estimate", "P": "Current estimate"}
        for c, label in labels.items():
            f = cell("%s7" % c).getFormula()
            if f != '=(Cfg_BaseYear+0)&" %s"' % label:
                raise RuntimeError("Expenditures %s7: unexpected heading %r" % (c, f))
            cell("%s7" % c).setFormula('=Cfg_BaseYear&CHAR(10)&"%s"' % label)

        # Column C wraps; the item rows keep an automatic height (design 17.3). The hidden
        # columns' text no longer wraps: the adjustment labels (J) made one-line rows two lines
        # tall.
        if not cell("C%d" % LAST_ROW).getString().startswith("▼ End of items"):
            raise RuntimeError("Expenditures C%d is not the end-of-items marker" % LAST_ROW)
        cell("C%d:C%d" % (FIRST_ROW, LAST_ROW - 1)).IsTextWrapped = True  # not the marker
        hidden = [c for c in "EFGHJ" if not sheet.Columns.getByName(c).IsVisible]
        if hidden != list("EFGHJ"):
            raise RuntimeError("Expenditures: hidden columns %s, not E-H and J" % hidden)
        for c in hidden:
            cell("%s%d:%s%d" % (c, FIRST_ROW, c, LAST_ROW)).IsTextWrapped = False

        self.fit_columns(sheet, ["L7:P7"], uniform=True)
        self.fit_columns(sheet, ["A7", "B%d:B%d" % (FIRST_ROW, LAST_ROW),
                                 "D%d:D%d" % (FIRST_ROW, LAST_ROW)])
        self.fit_columns(sheet, ["V3:AY5"], uniform=True)
        self.fit_fixed_row(sheet, 7)

    # -- Projected_Balances -------------------------------------------------------------------

    def projected(self):
        sheet = self.doc.Sheets.getByName("Projected_Balances")
        cell = sheet.getCellRangeByName
        if cell("C13").getString() != "Expected costs" or \
                cell("I13").getString() != "Diff from RFS_cashflow balance":
            raise RuntimeError("Projected_Balances C13, I13 are not the flagged columns")
        # The alarm brings its own number format: the columns' [RED] would override its white
        # text (design 17.4).
        for col in ("C", "I"):
            rng = cell("%s14:%s63" % (col, col))
            cf = rng.ConditionalFormat
            if cf.Count != 1:
                raise RuntimeError("Projected_Balances %s14:%s63: %d conditions, not 1"
                                   % (col, col, cf.Count))
            formula = cf.getByIndex(0).Formula1
            clear_conditions(rng)
            add_condition(rng, ALARM_NUMBER, formula, "FORMULA")
        # The difference from RFS_cashflow is rounded to whole dollars, the check's own
        # resolution: a remainder of 1e-8 showed as a red "-0" (design 17.4).
        for r in range(14, 64):
            c = cell("I%d" % r)
            want = "=$H{r}-INDEX(RFS_Cf_Balance;MATCH($A{r};RFS_Cf_Year;0))".format(r=r)
            if c.getFormula() != want:
                raise RuntimeError("Projected_Balances I%d: unexpected formula %r" % (r, c.getFormula()))
            c.setFormula("=ROUND(" + want[1:] + ";0)")
        self.fit_columns(sheet, ["D13:E13", "I13", "M13", "P13:R13", "Z13", "F4", "V4"])
        self.fit_span(sheet, "B6", "B", ["C"])
        self.fit_span(sheet, "R6", "R", ["S"])
        self.fit_fixed_row(sheet, 13)

    # -- other tabs ---------------------------------------------------------------------------

    def mtc(self):
        sheet = self.doc.Sheets.getByName("Medium_Term_Cashflow")
        self.fit_columns(sheet, ["A5:A25"])

    def schedule(self):
        sheet = self.doc.Sheets.getByName("Schedule_of_Investments")
        self.fit_columns(sheet, ["A3:J4", "A7:J40"])

    def settings(self):
        sheet = self.doc.Sheets.getByName("Settings")
        self.fit_columns(sheet, ["A4:A10"])

    def vendor_tabs(self):
        """The description columns wrap, and the item rows have an automatic height: no fixed
        heights on a data source (design 17.6)."""
        for name, columns, first in (("RFS_info", ("C", "D"), 10), ("RFS_fvalues", ("B",), 15)):
            sheet = self.doc.Sheets.getByName(name)
            cur = sheet.createCursor()
            cur.gotoEndOfUsedArea(False)
            last = cur.RangeAddress.EndRow + 1
            for c in columns:
                sheet.getCellRangeByName("%s%d:%s%d" % (c, first, c, last)).IsTextWrapped = True
            sheet.getCellRangeByName("A%d:A%d" % (first, last)).Rows.OptimalHeight = True

    def input_rows(self):
        """Every row with an input cell has an automatic height, so that it grows as the user
        types, also on a protected tab (design 17.3); the rows of Expenditures column C too."""
        for name in SPATAC_TABS:
            sheet = self.doc.Sheets.getByName(name)
            cur = sheet.createCursor()
            cur.gotoEndOfUsedArea(False)
            a = cur.RangeAddress
            area = sheet.getCellRangeByPosition(0, 0, a.EndColumn, a.EndRow)
            rows = set()
            for rng in area.getCellFormatRanges():
                if not rng.CellProtection.IsLocked:
                    ra = rng.RangeAddress
                    rows.update(range(ra.StartRow, ra.EndRow + 1))
            if name == "Expenditures":
                rows.update(range(FIRST_ROW - 1, LAST_ROW))
            fixed = [r for r in sorted(rows) if not sheet.Rows.getByIndex(r).OptimalHeight]
            for r in sorted(rows):
                sheet.Rows.getByIndex(r).OptimalHeight = True
            if fixed:
                self.report.append("%s: %d input rows made automatic" % (name, len(fixed)))

    def protect(self):
        sheets = self.doc.Sheets
        for name in SPATAC_TABS + VENDOR_TABS:
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
