"""Restructures the Spatac workbook as specified in docs/archive/design.md (phase 3).

Usage (LibreOffice's bundled Python carries the `uno` module):

    /opt/libreoffice26.2/program/python tools/migrate/migrate.py SOURCE.xlsx OUTPUT.xlsx

SOURCE is the pre-restructure workbook (Spatac-20261121.xlsx). It is opened read-only; nothing is
written to it. OUTPUT is a new workbook built from scratch, with the vendor RFS tabs copied across
and every Spatac tab rebuilt with defined names, input styles and formulas that follow the design.

The migration reproduces today's results. Known bugs are preserved here and fixed separately
(phase 4, `--fix` options).
"""

import argparse
import sys

import uno  # noqa: F401  (registers the com.sun.star import hook)
from com.sun.star.table import CellRangeAddress
from com.sun.star.util import CellProtection

from uno_session import Session

BASE_YEAR = 2026
RFS_YEARS = 28  # base year .. base year + 27, the span the old fvalues_spatac covered
HORIZON = 6  # Medium_Term_Cashflow years
SCHEDULE_YEARS = 11  # Schedule_of_Investments year columns: base .. base + 10
HELD_ROWS = 40
PLANNED_ROWS = 30

# Cell flags for clearContents.
CLEAR_CONTENT = 1 | 2 | 4 | 16  # values, dates, strings, formulas

INPUT_COLOR = 0xFFF2CC
CALC_COLOR = 0xE2EFDA
HEADER_COLOR = 0xD9D9D9
SEPARATOR_COLOR = 0xBFBFBF

MONEY = '#,##0;-#,##0;""'
MONEY_ZERO = "#,##0;-#,##0;0"
PERCENT = "0.00%"
DATE = "YYYY-MM-DD"
YEAR = "0"


# --------------------------------------------------------------------------------------------
# Helpers


def col(n):
    """1 -> A."""
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


def is_blank(v):
    return v is None or v == ""


def num(v):
    if isinstance(v, float):
        return v
    if isinstance(v, str):
        try:
            return float(v)
        except ValueError:
            return 0.0
    return 0.0


def lo_text(v):
    """What `v&""` gives in a formula: numbers as the shortest decimal, integers without '.0'."""
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        return repr(v)
    return v or ""


def trim(s):
    """Calc's TRIM: strips the ends and collapses runs of spaces."""
    return " ".join(part for part in str(s).split(" ") if part != "")


def key_of(item, text):
    return trim(lo_text(item)) + "|" + trim(lo_text(text))


class Builder:
    """Wraps a sheet with the few operations the build needs."""

    def __init__(self, doc, sheet):
        self.doc = doc
        self.sheet = sheet
        self.formats = doc.NumberFormats
        self.locale = doc.CharLocale

    def cell(self, a):
        return self.sheet.getCellRangeByName(a)

    def value(self, a, v):
        c = self.cell(a)
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            c.setValue(float(v))
        elif v is None or v == "":
            pass
        else:
            c.setString(str(v))
        return c

    def formula(self, a, f):
        c = self.cell(a)
        c.setFormula(f)
        return c

    def values(self, rng, rows):
        """Writes a block of constants; text stays text, so item numbers like 3.10 survive."""
        data = tuple(tuple("" if v is None else (float(v) if isinstance(v, (int, float)) else v) for v in row) for row in rows)
        self.sheet.getCellRangeByName(rng).setDataArray(data)

    def formulas(self, rng, rows):
        self.sheet.getCellRangeByName(rng).setFormulaArray(tuple(tuple(r) for r in rows))

    def fmt(self, rng, code):
        key = self.formats.queryKey(code, self.locale, False)
        if key == -1:
            key = self.formats.addNew(code, self.locale)
        self.sheet.getCellRangeByName(rng).NumberFormat = key

    def style(self, rng, name):
        self.sheet.getCellRangeByName(rng).CellStyle = name

    def bold(self, rng):
        self.sheet.getCellRangeByName(rng).CharWeight = 150.0

    def fill(self, rng, color):
        self.sheet.getCellRangeByName(rng).CellBackColor = color

    def wrap(self, rng):
        self.sheet.getCellRangeByName(rng).IsTextWrapped = True

    def width(self, letters, chars):
        # Roughly 0.19 cm per character of the default font.
        self.sheet.Columns.getByName(letters).Width = int(chars * 190)

    def height(self, row, points):
        self.sheet.Rows.getByIndex(row - 1).Height = int(points * 35.28)

    def group_cols(self, first, last, collapse=False):
        a = CellRangeAddress()
        a.Sheet = self.sheet.RangeAddress.Sheet
        a.StartColumn, a.EndColumn = colnum(first) - 1, colnum(last) - 1
        a.StartRow = a.EndRow = 0
        self.sheet.group(a, uno.Enum("com.sun.star.table.TableOrientation", "COLUMNS"))
        if collapse:
            self.sheet.hideDetail(a)

    def comment(self, a, text):
        addr = self.cell(a).CellAddress
        self.sheet.Annotations.insertNew(addr, text)


def define(doc, name, content):
    """A workbook-scoped name; `content` in Calc syntax, e.g. $Settings.$B$3."""
    names = doc.NamedRanges
    if names.hasByName(name):
        names.removeByName(name)
    origin = uno.createUnoStruct("com.sun.star.table.CellAddress")
    names.addNewByName(name, content, origin, 0)


def ref(sheet, rng):
    """Absolute Calc reference: ref('Settings', 'B3') -> $Settings.$B$3."""
    def absolute(part):
        letters = "".join(ch for ch in part if ch.isalpha())
        digits = "".join(ch for ch in part if ch.isdigit())
        return "$%s$%s" % (letters, digits) if digits else "$%s" % letters
    parts = rng.split(":")
    quoted = "'%s'" % sheet if not sheet.replace("_", "").isalnum() else sheet
    return "$%s.%s" % (quoted, ":".join(absolute(p) for p in parts))


def make_styles(doc):
    styles = doc.StyleFamilies.getByName("CellStyles")

    def add(name, color=None, locked=True, bold=False):
        if styles.hasByName(name):
            s = styles.getByName(name)
        else:
            s = doc.createInstance("com.sun.star.style.CellStyle")
            styles.insertByName(name, s)
        s.ParentStyle = "Default"
        if color is not None:
            s.CellBackColor = color
            s.IsCellBackgroundTransparent = False
        p = CellProtection()
        p.IsLocked = locked
        p.IsFormulaHidden = False
        p.IsHidden = False
        p.IsPrintHidden = False
        s.CellProtection = p
        if bold:
            s.CharWeight = 150.0
        return s

    add("Input", INPUT_COLOR, locked=False)
    add("Calculated", CALC_COLOR)
    add("Header", HEADER_COLOR, bold=True)
    title = add("Title", bold=True)
    title.CharHeight = 14.0


# --------------------------------------------------------------------------------------------
# Reading the source


def read_block(sheet, rng):
    return [list(r) for r in sheet.getCellRangeByName(rng).getDataArray()]


def read_annotations(sheet):
    out = []
    anns = sheet.Annotations
    for i in range(anns.Count):
        a = anns.getByIndex(i)
        out.append((a.Position.Row + 1, a.Position.Column + 1, a.String))
    return out


# --------------------------------------------------------------------------------------------
# The build


class Migration:
    def __init__(self, session, src_path, out_path, fixes):
        self.s = session
        self.src_path = src_path
        self.out_path = out_path
        self.fixes = fixes
        self.notes = []  # what the migration changed in the data, for the report

    def run(self):
        src = self.s.open(self.src_path)
        src.calculateAll()
        doc = self.s.new_calc()
        self.doc = doc
        make_styles(doc)

        self.import_vendor(src)
        self.read_source(src)
        self.build_settings()
        self.build_readme()
        self.build_expenditures()
        self.build_mtc()
        self.build_schedule()
        self.build_projected()
        self.finish_vendor_helpers()
        self.reparse()
        self.order_and_protect()

        doc.calculateAll()
        self.s.save_xlsx(doc, self.out_path)
        doc.close(True)
        src.close(True)
        return self.notes

    def reparse(self):
        """Formulas were written before every name they use existed (Expenditures reads
        Schedule_of_Investments, which reads Medium_Term_Cashflow, which reads both). Writing
        each formula again, now that all names are defined, resolves them."""
        for sheet in self.doc.Sheets:
            ranges = sheet.queryContentCells(16)  # formulas only
            for i in range(ranges.Count):
                rng = ranges.getByIndex(i)
                rng.setFormulaArray(rng.getFormulaArray())

    # -- vendor tabs ---------------------------------------------------------------------------

    def import_vendor(self, src):
        doc = self.doc
        sheets = doc.Sheets
        placeholder = sheets.getByIndex(0).Name
        for old, new in (("info", "RFS_info"), ("fvalues", "RFS_fvalues"), ("cashflow", "RFS_cashflow")):
            idx = sheets.importSheet(src, old, sheets.Count)
            sh = sheets.getByIndex(idx)
            sh.Name = new
        sheets.getByName(placeholder).Name = "README"

        fv = sheets.getByName("RFS_fvalues")
        # Invisible charts with broken series (Q15).
        for name in list(fv.Charts.ElementNames):
            fv.Charts.removeByName(name)
        # Stray unlabelled numbers in row 8 (Q15).
        fv.getCellRangeByName("X8:AZ8").clearContents(CLEAR_CONTENT)
        # Hidden rows 217-218 hold only zeros (Q15).
        fv.Rows.removeByIndex(216, 2)
        # Formatting out to the last column (Q15): keep through BA, then the helper column.
        self.trim_columns(fv, colnum("BA"))
        self.trim_columns(sheets.getByName("RFS_info"), colnum("H"))
        self.trim_columns(sheets.getByName("RFS_cashflow"), colnum("K"))
        self.notes.append(
            "RFS_fvalues: removed the invisible broken chart, the unlabelled numbers in X8:AZ8 "
            "and the hidden zero rows 217-218; formatting beyond the data removed on all RFS tabs."
        )

    def trim_columns(self, sheet, last_used):
        cols = sheet.Columns
        extra = cols.Count - last_used
        if extra > 0:
            cols.removeByIndex(last_used, extra)

    def finish_vendor_helpers(self):
        doc = self.doc
        info = doc.Sheets.getByName("RFS_info")
        b = Builder(doc, info)
        rows = 400
        b.formulas(
            "J1:K%d" % rows,
            [
                [
                    '=IF(OR($A{r}="";$B{r}="");"";IF(ISERROR(VALUE($B{r}&""));"";'
                    'TRIM($A{r}&"")&"|"&TEXT(VALUE($B{r}&"");"0000.0000")))'.format(r=r),
                    '=IF($J{r}="";"";IF(ISNA(MATCH($J{r};Exp_Key2;0));"No Expenditures row";""))'.format(r=r),
                ]
                for r in range(1, rows + 1)
            ],
        )
        b.style("J1:K%d" % rows, "Calculated")
        b.width("J", 18)
        b.width("K", 20)
        b.comment(
            "K1",
            "Columns J:K are not part of the vendor paste. J is each item's key (item number | "
            "GL account); K flags an RFS item with no row on Expenditures.",
        )
        define(doc, "RFS_Info_Key", ref("RFS_info", "J1:J%d" % rows))
        define(doc, "RFS_Info_Check", ref("RFS_info", "K1:K%d" % rows))

        fv = doc.Sheets.getByName("RFS_fvalues")
        b = Builder(doc, fv)
        b.formulas(
            "BC1:BC%d" % rows,
            [['=IF(AND($A{r}="";$B{r}="");"";TRIM($A{r}&"")&"|"&TRIM($B{r}&""))'.format(r=r)] for r in range(1, rows + 1)],
        )
        b.style("BC1:BC%d" % rows, "Calculated")
        b.width("BC", 30)
        b.comment(
            "BC1",
            "Column BC is not part of the vendor paste: each row's key (item number | "
            "description), which Expenditures uses to look up the RFS amounts.",
        )
        define(doc, "RFS_Fv_Key", ref("RFS_fvalues", "BC1:BC%d" % rows))
        define(doc, "RFS_Fv_Values", ref("RFS_fvalues", "C1:AZ%d" % rows))
        define(doc, "RFS_Fv_Years", ref("RFS_fvalues", "C9:AZ9"))
        define(doc, "RFS_Fv_Totals", ref("RFS_fvalues", "C10:AZ10"))

        for name, rng in (
            ("RFS_Cf_Year", "A1:A200"),
            ("RFS_Cf_Cost", "B1:B200"),
            ("RFS_Cf_Interest", "C1:C200"),
            ("RFS_Cf_Contrib", "D1:D200"),
            ("RFS_Cf_Balance", "G1:G200"),
            ("RFS_StudyYear", "A15"),
            ("RFS_Units", "C6"),
            ("RFS_MinBalance", "C10"),
            ("RFS_OpeningBalance", "C11"),
            ("RFS_PresentContribution", "C12"),
            ("RFS_ShortInterest", "G7"),
            ("RFS_LongInterest", "G8"),
            ("RFS_ShortInflation", "K7"),
            ("RFS_LongInflation", "K8"),
            ("RFS_CatchUpIncreases", "K9:K11"),
            ("RFS_SubsequentIncrease", "K12"),
        ):
            define(doc, name, ref("RFS_cashflow", rng))

    # -- source data -------------------------------------------------------------------------

    def read_source(self, src):
        sh = src.Sheets
        self.fs = read_block(sh.getByName("fvalues_spatac"), "A1:AP255")
        self.fs_comments = read_annotations(sh.getByName("fvalues_spatac"))
        self.mtc = read_block(sh.getByName("Medium_Term_Cashflow"), "A1:M69")
        self.inv = read_block(sh.getByName("Schedule_of_Investments"), "A1:T46")
        self.sim = read_block(sh.getByName("cashflow_spatac_simulator"), "A1:V65")
        self.sim_comments = read_annotations(sh.getByName("cashflow_spatac_simulator"))
        fv = read_block(sh.getByName("fvalues"), "A1:AZ218")
        # RFS amount per key and year, as the lookup will see it.
        years = {}
        for c in range(2, 52):
            y = fv[8][c]
            if isinstance(y, float):
                years[int(y)] = c
        self.rfs = {}
        for r in range(14, 216):
            k = key_of(fv[r][0], fv[r][1])
            self.rfs[k] = {y: num(fv[r][c]) for y, c in years.items()}
        info = read_block(sh.getByName("info"), "A1:H214")
        self.info_gl = {}
        for row in info:
            if isinstance(row[1], float):
                self.info_gl.setdefault(trim(lo_text(row[0])), []).append(row[1])

    def fs_cell(self, letters, row):
        return self.fs[row - 1][colnum(letters) - 1]

    # -- Settings ------------------------------------------------------------------------------

    def build_settings(self):
        doc = self.doc
        sh = doc.Sheets
        sh.insertNewByName("Settings", 1)
        b = Builder(doc, sh.getByName("Settings"))
        b.value("A1", "SETTINGS").CellStyle = "Title"
        b.value("A2", "Values used throughout the workbook. Yellow cells are inputs.")
        rows = [
            ("Base year (the current year; change it in January to roll the workbook forward)", BASE_YEAR, YEAR, "Cfg_BaseYear"),
            ("Minimum target cash balance (Board policy s.3.5), in dollars of the year below", 1000000, MONEY_ZERO, "Cfg_MinTarget"),
            ("Year in whose dollars the minimum target is stated", 2024, YEAR, "Cfg_MinTargetYear"),
            ("Assumed interest rate for planned new investments", 0.03, PERCENT, "Cfg_NewInvRate"),
        ]
        for i, (label, v, fmt, name) in enumerate(rows):
            r = 4 + i
            b.value("A%d" % r, label)
            b.value("B%d" % r, v)
            b.style("B%d" % r, "Input")
            b.fmt("B%d" % r, fmt)
            define(doc, name, ref("Settings", "B%d" % r))
        b.value("A9", "The minimum target is inflated from its year at the RFS short-term inflation rate.")
        b.width("A", 75)
        b.width("B", 14)
        self.notes.append(
            "Settings: the minimum target, previously hidden in the constant 1,040,400 on "
            "Medium_Term_Cashflow B24, is now $1,000,000 in 2024 dollars; the 3% rate for planned "
            "investments moved here from Medium_Term_Cashflow J13."
        )

    # -- README --------------------------------------------------------------------------------

    def build_readme(self):
        doc = self.doc
        b = Builder(doc, doc.Sheets.getByName("README"))
        lines = [
            ("SPATAC WORKBOOK", "Title"),
            ("Reserve fund project tracking and investment planning for the Antiquary (MTCC No. 1360).", None),
            ("Documentation: https://pvillela.github.io/spatac-workbook/ — what each tab is for, how to use it, and the January and new-RFS procedures.", None),
            ("", None),
            ("Tabs", "Header"),
            ("Settings — the base year, the minimum target cash balance and the rate assumed for planned investments.", None),
            ("Expenditures — every RFS component, with Spatac's approved amounts, actuals, estimates and reasons; the RFS amounts by year; and the Revised amounts the plan uses.", None),
            ("Medium_Term_Cashflow — the Investment Plan: six years of cash, expenditures, maturities and planned new investments against the minimum target.", None),
            ("Schedule_of_Investments — investments held and planned, with each one's cash flows by year.", None),
            ("Projected_Balances — the reserve fund balance projected to the end of the RFS, with Spatac's adjustments and what-if contribution increases.", None),
            ("RFS_info, RFS_fvalues, RFS_cashflow — the vendor's RFS spreadsheet, pasted in unchanged. Helper columns to the right are not part of the paste.", None),
            ("", None),
            ("Cell colours", "Header"),
            ("Yellow — an input: type here.", "Input"),
            ("Green — calculated next to inputs: read only.", "Calculated"),
            ("Everything else is calculated or a label. Sheets are protected without a password to prevent accidents; Review > Unprotect Sheet when you need to.", None),
            ("", None),
            ("Estimates", "Header"),
            ("Mid-year, Latest and Late-year estimates are entered only for items whose estimate changed: type the new full amount, never a difference. Blank means no change; the Current estimate column shows the result for every item.", None),
            ("", None),
            ("Excel users: enter formulas normally (no Ctrl+Shift+Enter) and do not use formulas that spill; LibreOffice cannot keep them.", None),
        ]
        for i, (text, style) in enumerate(lines):
            c = b.value("A%d" % (i + 1), text)
            if style:
                c.CellStyle = style
        b.width("A", 120)
        b.wrap("A1:A%d" % len(lines))

    # -- Expenditures --------------------------------------------------------------------------

    def build_expenditures(self):
        doc = self.doc
        doc.Sheets.insertNewByName("Expenditures", 2)
        sheet = doc.Sheets.getByName("Expenditures")
        b = Builder(doc, sheet)
        first, last = 7, 250  # data rows (old 9..252)
        rfs0, rev0 = colnum("W"), colnum("AY")
        rfs_cols = [col(rfs0 + k) for k in range(RFS_YEARS)]
        rev_cols = [col(rev0 + k) for k in range(RFS_YEARS)]

        # Titles and group labels.
        b.value("A1", "SPATAC FUTURE EXPENDITURE TABLE").CellStyle = "Title"
        b.formula("J2", '=Cfg_BaseYear&" — current year"')
        b.formula("O2", '=(Cfg_BaseYear+1)&" — next year"')
        b.formula("R2", '=(Cfg_BaseYear+2)&" — year after"')
        b.value("W2", "RFS amounts by year (from RFS_fvalues)")
        b.value("AY2", "Revised amounts by year (what the plan uses)")
        b.bold("A2:BZ2")

        # Row 3: vendor totals (check); row 4: totals; row 5: provenance notes.
        b.value("C3", "RFS total per vendor")
        b.value("C4", "Total")
        b.value("C5", "As of / source (edit when you update a column)")
        for k, c in enumerate(rfs_cols):
            b.formula(
                "%s3" % c,
                "=IFERROR(-INDEX(RFS_Fv_Totals;1;MATCH(%s$6;RFS_Fv_Years;0));0)" % c,
            )
        for c in ["J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"] + rfs_cols + rev_cols:
            b.formula("%s4" % c, "=SUM(%s%d:%s%d)" % (c, first, c, last))
        b.fmt("J3:BZ4", MONEY_ZERO)
        b.bold("A4:BZ4")
        provenance = {
            "K": "to Oct 26",
            "L": "per KB in July adj by BM",
            "M": "from KB",
            "O": "2026-07-08",
            "P": "2026-11-04",
            "R": "2026-07-08",
            "S": "2026-11-04",
        }
        for c, text in provenance.items():
            b.value("%s5" % c, text)
            b.style("%s5" % c, "Input")
        b.wrap("A5:V5")

        # Row 6: headers.
        headers = {
            "A": "Row #",
            "B": "Item",
            "C": "Building element",
            "D": "GL account",
            "E": "GL account name",
            "F": "Spatac row",
            "G": "Key (item|element)",
            "H": "Key (item|GL)",
            "I": "RFS check",
            "U": "Reason for change",
            "V": "Working notes",
        }
        for c, t in headers.items():
            b.value("%s6" % c, t)
        year_headers = {
            "J": "Approved",
            "K": "Actual to date",
            "L": "Mid-year estimate",
            "M": "Latest estimate",
            "N": "Current estimate",
            "O": "Mid-year estimate",
            "P": "Late-year estimate",
            "Q": "Current estimate",
            "R": "Mid-year estimate",
            "S": "Late-year estimate",
            "T": "Current estimate",
        }
        offset = {"J": 0, "K": 0, "L": 0, "M": 0, "N": 0, "O": 1, "P": 1, "Q": 1, "R": 2, "S": 2, "T": 2}
        for c, t in year_headers.items():
            b.formula("%s6" % c, '=(Cfg_BaseYear+%d)&" %s"' % (offset[c], t))
        b.formula("W6", "=Cfg_BaseYear")
        b.formula("AY6", "=Cfg_BaseYear")
        for k in range(1, RFS_YEARS):
            b.formula("%s6" % rfs_cols[k], "=%s6+1" % rfs_cols[k - 1])
            b.formula("%s6" % rev_cols[k], "=%s6+1" % rev_cols[k - 1])
        b.style("A6:BZ6", "Header")
        b.wrap("A6:BZ6")
        b.height(6, 42)

        # Data rows.
        identity, inputs, notes = [], [], []
        zero_rule = {"M": 0, "P": 0, "S": 0}
        spatac_rows = 0
        for r in range(first, last + 1):
            old = r + 2
            item = self.fs_cell("B", old)
            element = self.fs_cell("C", old)
            gl = self.fs_cell("D", old)
            if old == 187 and lo_text(item) == "16" and trim(element) == "Garbage Compactor":
                item = "16.1"
                self.notes.append(
                    "Expenditures row %d (old fvalues_spatac row 187): item number 16 corrected "
                    "to 16.1, the RFS item it carries (Garbage Compactor, $28,154 in 2030)." % r
                )
            key = key_of(item, element)
            in_rfs = key in self.rfs
            if in_rfs and is_blank(gl):
                candidates = self.info_gl.get(trim(lo_text(item)), [])
                if len(candidates) == 1:
                    gl = "%.4f" % candidates[0]
                    self.notes.append(
                        "Expenditures row %d (item %s): blank GL account filled from RFS_info "
                        "(%s)." % (r, lo_text(item), gl)
                    )
            spatac = (not in_rfs) and not is_blank(gl)
            spatac_rows += spatac
            if isinstance(gl, float):
                gl = "%.4f" % gl
            identity.append([self.fs_cell("A", old), item, element, gl, self.fs_cell("E", old), "S" if spatac else ""])

            rfs = self.rfs.get(key, {})
            approved, actual, mid, latest = (self.fs_cell(c, old) for c in "IJKL")
            mid1, late1, mid2, late2 = (self.fs_cell(c, old) for c in "NPOQ")
            # Blank used to mean zero; it now means "no change", so a zero is written wherever
            # the fallback would otherwise produce a different amount.
            if is_blank(latest) and (num(mid) != 0 or num(approved) != 0):
                latest = 0.0
                zero_rule["M"] += 1
            if is_blank(late1) and (num(mid1) != 0 or rfs.get(BASE_YEAR + 1, 0) != 0):
                late1 = 0.0
                zero_rule["P"] += 1
            if is_blank(late2) and (num(mid2) != 0 or rfs.get(BASE_YEAR + 2, 0) != 0):
                late2 = 0.0
                zero_rule["S"] += 1
            inputs.append([approved, actual, mid, latest])
            notes.append((mid1, late1, mid2, late2, self.fs_cell("M", old)))

        b.values("A%d:F%d" % (first, last), identity)
        b.values("J%d:M%d" % (first, last), inputs)
        b.values("O%d:P%d" % (first, last), [[n[0], n[1]] for n in notes])
        b.values("R%d:S%d" % (first, last), [[n[2], n[3]] for n in notes])
        b.values("V%d:V%d" % (first, last), [[n[4]] for n in notes])
        self.notes.append(
            "Expenditures: %d rows marked as Spatac rows (no RFS item). Explicit zeros written "
            "where a blank used to mean zero: %d Latest, %d next-year Late-year, %d year-after "
            "Late-year estimates. Formulas in old column L (=MAX(I,J), =J) and in N30, O70, Q70, "
            "O208 became their values." % (spatac_rows, zero_rule["M"], zero_rule["P"], zero_rule["S"])
        )

        formulas = []
        for r in range(first, last + 1):
            row = [
                '=TRIM($B{r}&"")&"|"&TRIM($C{r}&"")',
                '=TRIM($B{r}&"")&"|"&IFERROR(TEXT(VALUE($D{r}&"");"0000.0000");"")',
                '=IF(OR($D{r}="";$F{r}<>"");"";IF(ISNA(MATCH($H{r};RFS_Info_Key;0));"Not in RFS";'
                'IF(ISNA(MATCH($G{r};RFS_Fv_Key;0));"No RFS amounts";"")))',
            ]
            formulas.append([f.format(r=r) for f in row])
        b.formulas("G%d:I%d" % (first, last), formulas)
        b.formulas(
            "N%d:N%d" % (first, last),
            [['=IF($M{r}<>"";$M{r};IF($L{r}<>"";$L{r};N($J{r})))'.format(r=r)] for r in range(first, last + 1)],
        )
        b.formulas(
            "Q%d:Q%d" % (first, last),
            [['=IF($P{r}<>"";$P{r};IF($O{r}<>"";$O{r};$X{r}))'.format(r=r)] for r in range(first, last + 1)],
        )
        b.formulas(
            "T%d:T%d" % (first, last),
            [['=IF($S{r}<>"";$S{r};IF($R{r}<>"";$R{r};$Y{r}))'.format(r=r)] for r in range(first, last + 1)],
        )
        b.formulas(
            "%s%d:%s%d" % (rfs_cols[0], first, rfs_cols[-1], last),
            [
                [
                    "=IFERROR(INDEX(RFS_Fv_Values;MATCH($G{r};RFS_Fv_Key;0);MATCH({c}$6;RFS_Fv_Years;0));0)".format(r=r, c=c)
                    for c in rfs_cols
                ]
                for r in range(first, last + 1)
            ],
        )
        b.formulas(
            "%s%d:%s%d" % (rev_cols[0], first, rev_cols[-1], last),
            [
                [
                    "=IF({c}$6=Cfg_BaseYear;$N{r};IF({c}$6=Cfg_BaseYear+1;$Q{r};IF({c}$6=Cfg_BaseYear+2;$T{r};{f}{r})))".format(
                        r=r, c=c, f=f
                    )
                    for c, f in zip(rev_cols, rfs_cols)
                ]
                for r in range(first, last + 1)
            ],
        )

        # Styles and formats.
        b.style("A%d:F%d" % (first, last), "Input")
        b.style("J%d:M%d" % (first, last), "Input")
        b.style("O%d:P%d" % (first, last), "Input")
        b.style("R%d:S%d" % (first, last), "Input")
        b.style("U%d:V%d" % (first, last), "Input")
        b.style("G%d:I%d" % (first, last), "Calculated")
        for c in "NQT":
            b.style("%s%d:%s%d" % (c, first, c, last), "Calculated")
        b.fmt("J%d:T%d" % (first, last), MONEY)
        b.fmt("%s%d:%s%d" % (rfs_cols[0], first, rev_cols[-1], last), MONEY)
        b.fmt("W6:BZ6", YEAR)
        widths = {"A": 6, "B": 8, "C": 34, "D": 10, "E": 24, "F": 6, "G": 20, "H": 16, "I": 12, "U": 30, "V": 40}
        for c in "JKLMNOPQRST":
            widths[c] = 11
        for c in rfs_cols + rev_cols:
            widths[c] = 10
        for c, w in widths.items():
            b.width(c, w)
        b.wrap("U%d:V%d" % (first, last))

        # Groups (Q46): row number and GL account name, helper keys, each year's inputs, and the
        # two year blocks.
        b.group_cols("A", "A", collapse=True)
        b.group_cols("E", "E", collapse=True)
        b.group_cols("G", "H", collapse=True)
        b.group_cols("J", "M")
        b.group_cols("O", "P")
        b.group_cols("R", "S")
        b.group_cols(rfs_cols[0], rfs_cols[-1])
        b.group_cols(rev_cols[0], rev_cols[-1])

        # Comments, moved with their cells.
        colmap = {"A": "A", "B": "B", "C": "C", "D": "D", "E": "E", "H": "W", "I": "J", "J": "K", "K": "L", "L": "M", "M": "V", "N": "O", "O": "R", "P": "P", "Q": "S"}
        for row, c, text in self.fs_comments:
            letters = col(c)
            target = colmap.get(letters)
            new_row = 6 if row <= 8 else row - 2
            if target is None:
                target = "V"
                text = "[Comment from old column %s] %s" % (letters, text)
            b.comment("%s%d" % (target, new_row), text)

        # Names.
        def name(n, letters_from, letters_to=None):
            letters_to = letters_to or letters_from
            define(doc, n, ref("Expenditures", "%s%d:%s%d" % (letters_from, first, letters_to, last)))

        for n, c in (
            ("Exp_RowNumber", "A"), ("Exp_Item", "B"), ("Exp_Element", "C"), ("Exp_GL", "D"),
            ("Exp_GLName", "E"), ("Exp_SpatacRow", "F"), ("Exp_Key", "G"), ("Exp_Key2", "H"),
            ("Exp_Check", "I"), ("Exp_Approved", "J"), ("Exp_Actual", "K"), ("Exp_MidEst", "L"),
            ("Exp_Latest", "M"), ("Exp_Current", "N"), ("Exp_MidEst1", "O"), ("Exp_LateEst1", "P"),
            ("Exp_Current1", "Q"), ("Exp_MidEst2", "R"), ("Exp_LateEst2", "S"), ("Exp_Current2", "T"),
            ("Exp_Reason", "U"), ("Exp_Notes", "V"),
        ):
            name(n, c)
        name("Exp_RFS", rfs_cols[0], rfs_cols[-1])
        name("Exp_Revised", rev_cols[0], rev_cols[-1])
        define(doc, "Exp_RFSYears", ref("Expenditures", "%s6:%s6" % (rfs_cols[0], rfs_cols[-1])))
        define(doc, "Exp_RevisedYears", ref("Expenditures", "%s6:%s6" % (rev_cols[0], rev_cols[-1])))

        self.freeze(sheet, colnum("F") - 1, 6)

    def freeze(self, sheet, cols, rows):
        ctrl = self.doc.CurrentController
        ctrl.setActiveSheet(sheet)
        ctrl.freezeAtPosition(cols, rows)

    # -- Medium_Term_Cashflow ------------------------------------------------------------------

    def mtc_cell(self, letters, row):
        return self.mtc[row - 1][colnum(letters) - 1]

    def build_mtc(self):
        doc = self.doc
        doc.Sheets.insertNewByName("Medium_Term_Cashflow", 3)
        sheet = doc.Sheets.getByName("Medium_Term_Cashflow")
        b = Builder(doc, sheet)
        yc = [col(2 + k) for k in range(HORIZON)]  # B..G

        b.value("A1", self.mtc_cell("A", 1)).CellStyle = "Title"
        b.style("A1", "Input")
        b.bold("A1")
        b.formula("B3", "=Cfg_BaseYear")
        for k in range(1, HORIZON):
            b.formula("%s3" % yc[k], "=%s3+1" % yc[k - 1])
        b.style("A3:G3", "Header")
        b.fmt("B3:G3", YEAR)

        # Phase 4 fix: the carry-over sits below the closing balance, then the adjustment.
        fixed = "closing" in self.fixes
        carry, adj, minimum = (23, 24, 25) if fixed else (19, 22, 24)
        labels = {
            5: self.mtc_cell("A", 5),
            6: self.mtc_cell("A", 6),
            8: "Planned contributions (RFS)",
            9: "Revised expenditures (from Expenditures)",
            10: "Held investments maturing + interest payments (from Schedule_of_Investments)",
            11: "Planned investments maturing (from the grid below)",
            12: "Planned new investments",
            13: "1-year maturity",
            14: "2-year maturity",
            15: "3-year maturity",
            16: "4-year maturity",
            17: "5-year maturity (balancing figure)",
            carry: "Additional cash carry-over",
            21: "Closing cash balance",
            adj: "Adjustment to minimum cash balance" if fixed else "Adjustment to cash balance",
            minimum: "Minimum closing cash balance (Settings, inflated)",
            28: "RFS expenditures",
            29: "Revised expenditures",
            30: "Variance",
            32: "RFS investment balance",
            33: "Projected investment balance",
            34: "Variance",
        }
        for r, t in labels.items():
            b.value("A%d" % r, t)

        b.value("B5", self.mtc_cell("B", 5))
        b.style("B5", "Input")
        for k in range(1, HORIZON):
            b.formula("%s5" % yc[k], "=%s21" % yc[k - 1])
        for c in yc:
            b.value("%s6" % c, self.mtc_cell(c, 6))
            b.style("%s6" % c, "Input")
            b.formula("%s8" % c, "=INDEX(RFS_Cf_Contrib;MATCH(%s$3;RFS_Cf_Year;0))" % c)
            b.formula("%s9" % c, "=-SUM(INDEX(Exp_Revised;0;MATCH(%s$3;Exp_RevisedYears;0)))" % c)
            b.formula("%s10" % c, "=SUM(INDEX(Inv_HeldFlows;0;MATCH(%s$3;Inv_Years;0)))" % c)
            # Planned investments maturing this year, from the grid columns of earlier years only:
            # reading the whole grid would make the 5-year balancing row depend on itself.
            if c == yc[0]:
                b.value("%s11" % c, 0.0)
            else:
                prev = yc[yc.index(c) - 1]
                b.formula(
                    "%s11" % c,
                    "=SUMPRODUCT($B$13:{p}17*((COLUMN($B$13:{p}17)-COLUMN($B$13))+(ROW($B$13:{p}17)-ROW($B$13)+1)"
                    "=(COLUMN({c}13)-COLUMN($B$13)))*(1+Cfg_NewInvRate)^(ROW($B$13:{p}17)-ROW($B$13)+1))".format(p=prev, c=c),
                )
            b.formula("%s12" % c, "=-SUM(%s13:%s17)" % (c, c))
            for r in (13, 14, 15, 16):
                b.value("%s%d" % (c, r), self.mtc_cell(c, r))
                b.style("%s%d" % (c, r), "Input")
            b.formula(
                "%s17" % c,
                "=SUM({c}4:{c}11)-SUM({c}13:{c}16)-{c}{k}-{c}{m}-{c}{a}".format(
                    c=c, k=carry, m=minimum, a=adj
                ),
            )
            b.value("%s%d" % (c, carry), self.mtc_cell(c, 19))
            b.style("%s%d" % (c, carry), "Input")
            # Phase 4 fix: the balancing figure already leaves the carry-over and the adjustment
            # in cash, so the closing balance adds nothing more. The old formula added the
            # adjustment a second time.
            if fixed:
                b.formula("%s21" % c, "=SUM({c}5:{c}12)".format(c=c))
            else:
                b.formula("%s21" % c, "=SUM({c}5:{c}12)+{c}22".format(c=c))
            b.value("%s%d" % (c, adj), self.mtc_cell(c, 22))
            b.style("%s%d" % (c, adj), "Input")
            b.formula(
                "%s%d" % (c, minimum),
                "=Cfg_MinTarget*(1+RFS_ShortInflation)^(%s$3-Cfg_MinTargetYear)" % c,
            )
            b.formula("%s28" % c, "=-SUM(INDEX(Exp_RFS;0;MATCH(%s$3;Exp_RFSYears;0)))" % c)
            b.formula("%s29" % c, "=%s9" % c)
            b.formula("%s30" % c, "=%s29-%s28" % (c, c))
            b.formula("%s32" % c, "=INDEX(PB_RFSBalance;MATCH(%s$3;PB_Year;0))" % c)
            b.formula("%s33" % c, "=INDEX(PB_Balance;MATCH(%s$3;PB_Year;0))" % c)
            b.formula("%s34" % c, "=%s33-%s32" % (c, c))
        b.fmt("B5:G34", MONEY_ZERO)
        for r in (21, minimum):
            b.bold("A%d:G%d" % (r, r))
        for r in (26, 36):
            b.fill("A%d:G%d" % (r, r), SEPARATOR_COLOR)
        b.value("I33", "Based on the assumptions in Projected_Balances;")
        b.value("I34", "go there for longer-term projections.")
        b.width("A", 62)
        for c in yc:
            b.width(c, 13)

        g = lambda first, last: ref("Medium_Term_Cashflow", "%s:%s" % (first, last))
        for n, rng in (
            ("MTC_Title", "A1:A1"),
            ("MTC_Years", "B3:G3"),
            ("MTC_OpeningCash", "B5:G5"),
            ("MTC_Payables", "B6:G6"),
            ("MTC_Contributions", "B8:G8"),
            ("MTC_Expenditures", "B9:G9"),
            ("MTC_Maturities", "B10:G10"),
            ("MTC_PlannedMaturities", "B11:G11"),
            ("MTC_NewInvTotal", "B12:G12"),
            ("MTC_NewInv", "B13:G17"),
            ("MTC_CarryOver", "B%d:G%d" % (carry, carry)),
            ("MTC_ClosingCash", "B21:G21"),
            ("MTC_Adjustment", "B%d:G%d" % (adj, adj)),
            ("MTC_MinCash", "B%d:G%d" % (minimum, minimum)),
            ("MTC_RFSExpenditures", "B28:G28"),
            ("MTC_RevisedExpenditures", "B29:G29"),
            ("MTC_ExpVariance", "B30:G30"),
            ("MTC_RFSBalance", "B32:G32"),
            ("MTC_ProjectedBalance", "B33:G33"),
            ("MTC_BalanceVariance", "B34:G34"),
            ("MTC_CashSection", "A3:G%d" % minimum),
            ("MTC_ComparisonSection", "A28:G30"),
        ):
            a, z = rng.split(":")
            define(doc, n, g(a, z))
        self.freeze(sheet, 1, 3)
        self.notes.append(
            "Medium_Term_Cashflow: maturities and interest now come from Schedule_of_Investments "
            "year columns instead of hand-placed cells (old rows 38-67); the scratch area J37:K41 "
            "and the hand-built maturity row 41 are gone; rows 32-33 read Projected_Balances "
            "instead of TRANSPOSE arrays."
        )

    # -- Schedule_of_Investments ---------------------------------------------------------------

    def inv_cell(self, letters, row):
        return self.inv[row - 1][colnum(letters) - 1]

    def build_schedule(self):
        doc = self.doc
        doc.Sheets.insertNewByName("Schedule_of_Investments", 4)
        sheet = doc.Sheets.getByName("Schedule_of_Investments")
        b = Builder(doc, sheet)
        yc = [col(colnum("N") + k) for k in range(SCHEDULE_YEARS)]  # N..X
        held0, held1 = 7, 7 + HELD_ROWS - 1  # 7..46
        plan_hdr = held1 + 2  # 48
        plan0, plan1 = plan_hdr + 1, plan_hdr + PLANNED_ROWS  # 49..78
        total = plan1 + 2  # 80

        b.value("A1", self.inv_cell("A", 1)).CellStyle = "Input"
        b.bold("A1")
        heads3 = ["Name", "Due", "Purchased", "Book", "Interest", "Term", "Compounding", "Annual", "Maturity", "Last-year", "Purchase", "Due", "Grid"]
        heads4 = ["", "date", "date", "value", "rate", "(years)", "(A or C)", "interest", "value", "cash flow", "year", "year", "#"]
        for k, (h3, h4) in enumerate(zip(heads3, heads4)):
            c = col(k + 1)
            b.value("%s3" % c, h3)
            b.value("%s4" % c, h4)
        b.value("N3", "Cash flows by year: interest, then principal + interest at maturity")
        b.formula("N5", "=Cfg_BaseYear")
        for k in range(1, SCHEDULE_YEARS):
            b.formula("%s5" % yc[k], "=%s5+1" % yc[k - 1])
        b.style("A3:X5", "Header")
        b.fmt("N5:X5", YEAR)
        b.value("A6", "Held (type in each investment; delete it once it has matured)")
        b.bold("A6")

        # Held rows: the 25 current investments, then empty input rows.
        data = []
        for r in range(held0, held1 + 1):
            if r <= 31:
                data.append([self.inv_cell(c, r) for c in "ABCDE"] + [None, self.inv_cell("G", r)])
            else:
                data.append([None] * 7)
        b.values("A%d:E%d" % (held0, held1), [row[:5] for row in data])
        b.values("G%d:G%d" % (held0, held1), [[row[6]] for row in data])
        compound = "^5"  # phase 4 fix: the investment's own term
        if "compounding" in self.fixes:
            compound = "^$F{r}"
        held = []
        for r in range(held0, held1 + 1):
            held.append(
                [
                    '=IF(OR($B{r}="";$C{r}="");"";ROUND(DATEDIF($C{r};$B{r};"m")/12;0))',
                    '=IF($D{r}="";"";IF($G{r}="C";0;$D{r}*$E{r}))',
                    '=IF($D{r}="";"";IF($G{r}="C";$D{r}*(1+$E{r})' + compound + ";$D{r}))",
                    '=IF($D{r}="";"";$I{r}+$H{r})',
                    '=IF($C{r}="";"";YEAR($C{r}))',
                    '=IF($B{r}="";"";YEAR($B{r}))',
                ]
            )
            held[-1] = [f.format(r=r) for f in held[-1]]
        b.formulas("F%d:F%d" % (held0, held1), [[h[0]] for h in held])
        b.formulas("H%d:L%d" % (held0, held1), [h[1:] for h in held])
        b.style("A%d:E%d" % (held0, held1), "Input")
        b.style("G%d:G%d" % (held0, held1), "Input")

        # Planned rows, filled from the Medium_Term_Cashflow grid (term x purchase year).
        b.value("A%d" % plan_hdr, "Planned (filled automatically from the planned new investments on Medium_Term_Cashflow)")
        b.bold("A%d" % plan_hdr)
        planned = []
        for r in range(plan0, plan1 + 1):
            k = r - plan0 + 1
            planned.append(
                [
                    '=IF($M{r}="";"";"Planned "&$F{r}&"-year investment, bought "&$K{r})',  # A
                    '=IF($M{r}="";"";INDEX(MTC_NewInv;$F{r};INT(($M{r}-1)/ROWS(MTC_NewInv))+1))',  # D
                    '=IF($M{r}="";"";Cfg_NewInvRate)',  # E
                    '=IF($M{r}="";"";MOD($M{r}-1;ROWS(MTC_NewInv))+1)',  # F
                    '=IF($M{r}="";"";"C")',  # G
                    '=IF($M{r}="";"";0)',  # H
                    '=IF($M{r}="";"";$D{r}*(1+$E{r})^$F{r})',  # I
                    '=IF($M{r}="";"";$I{r}+$H{r})',  # J
                    '=IF($M{r}="";"";INDEX(MTC_Years;1;INT(($M{r}-1)/ROWS(MTC_NewInv))+1))',  # K
                    '=IF($M{r}="";"";$K{r}+$F{r})',  # L
                    "=IFERROR(AGGREGATE(15;6;((COLUMN(MTC_NewInv)-COLUMN(INDEX(MTC_NewInv;1;1)))*ROWS(MTC_NewInv)"
                    "+ROW(MTC_NewInv)-ROW(INDEX(MTC_NewInv;1;1))+1)/(MTC_NewInv<>0);%d);\"\")" % k,  # M
                ]
            )
            planned[-1] = [f.format(r=r) if "{r}" in f else f for f in planned[-1]]
        b.formulas("A%d:A%d" % (plan0, plan1), [[p[0]] for p in planned])
        b.formulas("D%d:L%d" % (plan0, plan1), [p[1:10] for p in planned])
        b.formulas("M%d:M%d" % (plan0, plan1), [[p[10]] for p in planned])
        b.style("A%d:M%d" % (plan0, plan1), "Calculated")

        # Year columns for both sections.
        rows = list(range(held0, held1 + 1)) + list(range(plan0, plan1 + 1))
        flows = '=IF($J{r}="";0;IF({c}$5=$L{r};$J{r};IF(AND({c}$5>$K{r};{c}$5<$L{r});$H{r};0)))'
        b.formulas("N%d:X%d" % (held0, held1), [[flows.format(r=r, c=c) for c in yc] for r in range(held0, held1 + 1)])
        b.formulas("N%d:X%d" % (plan0, plan1), [[flows.format(r=r, c=c) for c in yc] for r in range(plan0, plan1 + 1)])
        b.value("A%d" % total, "Total cash flows")
        for c in yc:
            b.formula("%s%d" % (c, total), "=SUM(%s%d:%s%d)" % (c, held0, c, plan1))
        b.bold("A%d:X%d" % (total, total))
        del rows

        b.fmt("B%d:C%d" % (held0, held1), DATE)
        b.fmt("D%d:D%d" % (held0, plan1), MONEY_ZERO)
        b.fmt("E%d:E%d" % (held0, plan1), PERCENT)
        b.fmt("H%d:J%d" % (held0, plan1), MONEY_ZERO)
        b.fmt("N%d:X%d" % (held0, total), MONEY)
        widths = {"A": 44, "B": 11, "C": 11, "D": 11, "E": 8, "F": 7, "G": 11, "H": 10, "I": 11, "J": 11, "K": 8, "L": 8, "M": 5}
        for c, w in widths.items():
            b.width(c, w)
        for c in yc:
            b.width(c, 10)
        b.group_cols("K", "M", collapse=True)

        span = lambda a, z: ref("Schedule_of_Investments", "%s:%s" % (a, z))
        for n, c in (
            ("Inv_Name", "A"), ("Inv_Due", "B"), ("Inv_Purchased", "C"), ("Inv_Amount", "D"),
            ("Inv_Rate", "E"), ("Inv_Term", "F"), ("Inv_Compounding", "G"),
            ("Inv_AnnualInterest", "H"), ("Inv_MaturityValue", "I"), ("Inv_LastYearCash", "J"),
            ("Inv_PurchaseYear", "K"), ("Inv_DueYear", "L"),
        ):
            define(doc, n, span("%s%d" % (c, held0), "%s%d" % (c, plan1)))
        define(doc, "Inv_Title", span("A1", "A1"))
        define(doc, "Inv_Held", span("A%d" % held0, "X%d" % held1))
        define(doc, "Inv_Planned", span("A%d" % plan0, "X%d" % plan1))
        define(doc, "Inv_Years", span("N5", "X5"))
        define(doc, "Inv_YearFlows", span("N%d" % held0, "X%d" % plan1))
        define(doc, "Inv_HeldFlows", span("N%d" % held0, "X%d" % held1))
        define(doc, "Inv_PlannedFlows", span("N%d" % plan0, "X%d" % plan1))
        define(doc, "Inv_Totals", span("N%d" % total, "X%d" % total))
        self.freeze(sheet, 1, 5)
        self.notes.append(
            "Schedule_of_Investments: removed columns L:R, the monthly cash diary (rows 37-46) "
            "and every #NAME? cell; added year columns and a Planned section filled from the "
            "Medium_Term_Cashflow grid."
        )

    # -- Projected_Balances --------------------------------------------------------------------

    def sim_cell(self, letters, row):
        return self.sim[row - 1][colnum(letters) - 1]

    def build_projected(self):
        doc = self.doc
        doc.Sheets.insertNewByName("Projected_Balances", 5)
        sheet = doc.Sheets.getByName("Projected_Balances")
        b = Builder(doc, sheet)
        first, n_years = 13, 50
        last = first + n_years - 1  # 62
        last30 = first + 29  # 2024..2053

        b.value("A1", "PROJECTED RESERVE FUND BALANCES").CellStyle = "Title"
        b.value("A2", "RFS figures from RFS_cashflow; Spatac adjustments are Revised minus RFS from Expenditures.")
        b.value("A4", "Contributions follow the RFS schedule through")
        b.value("B4", 2027.0)
        b.value("A5", "Contribution increase, 1st year after")
        b.value("A6", "Contribution increase, 2nd year after")
        b.value("A7", "Contribution increase, 3rd year after")
        for r, src in ((5, "Q9"), (6, "Q10"), (7, "Q11")):
            b.value("B%d" % r, num(self.sim_cell(src[0], int(src[1:]))))
        b.value("A8", "Subsequent annual increase")
        b.formula("B8", "=RFS_SubsequentIncrease")
        b.style("B4:B8", "Input")
        b.fmt("B4", YEAR)
        b.fmt("B5:B8", PERCENT)
        define(doc, "PB_ContribFixedThrough", ref("Projected_Balances", "B4"))
        define(doc, "PB_CatchUp", ref("Projected_Balances", "B5:B7"))
        define(doc, "PB_SubsequentIncrease", ref("Projected_Balances", "B8"))

        b.value("D4", "Historical inflation rate 2021–2024")
        b.value("E4", num(self.sim_cell("T", 4)))
        b.value("D5", "Unlabelled value carried over from the old simulator (Q7)")
        b.value("E5", num(self.sim_cell("Q", 7)))
        b.fmt("E4:E5", PERCENT)

        b.value("G4", "Lowest over 30 years")
        b.value("H4", "Amount")
        b.value("I4", "Year")
        lows = (
            ("Projected: margin over minimum", "P"),
            ("Projected: margin over inflated minimum", "Q"),
            ("Projected: margin over minimum (PV)", "R"),
            ("RFS: margin over minimum", "S"),
            ("RFS: margin over inflated minimum", "T"),
            ("RFS: margin over minimum (PV)", "U"),
        )
        for i, (label, c) in enumerate(lows):
            r = 5 + i
            b.value("G%d" % r, label)
            b.formula("H%d" % r, "=MIN(%s%d:%s%d)" % (c, first, c, last30))
            b.formula("I%d" % r, "=INDEX($A$%d:$A$%d;MATCH(H%d;%s%d:%s%d;0))" % (first, last30, r, c, first, c, last30))
        b.style("G4:I4", "Header")
        b.fmt("H5:H10", MONEY_ZERO)
        b.fmt("I5:I10", YEAR)
        define(doc, "PB_Lowest", ref("Projected_Balances", "G5:I10"))

        heads = [
            "Year", "RFS expected costs", "RFS interest", "RFS contributions", "RFS balance",
            "Actual expenditures", "Actual closing balance", "Spatac adjustment", "Adjusted costs",
            "Contribution increase", "Contributions", "Interest", "Projected balance",
            "Variance from RFS", "Inflator", "Margin over minimum", "Margin over inflated minimum",
            "Margin over minimum (PV)", "RFS margin over minimum", "RFS margin over inflated minimum",
            "RFS margin over minimum (PV)", "Contribution per unit per month", "",
            "Check: RFS interest recomputed", "Check: RFS contributions recomputed",
            "Check: RFS balance recomputed", "Check",
        ]
        for k, h in enumerate(heads):
            b.value("%s12" % col(k + 1), h)
        b.style("A12:AA12", "Header")
        b.wrap("A12:AA12")
        b.height(12, 42)

        rate = "IF($A{r}<=RFS_StudyYear+3;RFS_ShortInterest;RFS_LongInterest)"
        rows = []
        for r in range(first, last + 1):
            p = r - 1
            firstrow = r == first
            year = "=RFS_StudyYear" if firstrow else "=$A%d+1" % p
            row = [
                year,
                "=IFERROR(INDEX(RFS_Cf_Cost;MATCH($A{r};RFS_Cf_Year;0));0)",
                "=IFERROR(INDEX(RFS_Cf_Interest;MATCH($A{r};RFS_Cf_Year;0));0)",
                "=IFERROR(INDEX(RFS_Cf_Contrib;MATCH($A{r};RFS_Cf_Year;0));0)",
                "=IFERROR(INDEX(RFS_Cf_Balance;MATCH($A{r};RFS_Cf_Year;0));0)",
            ]
            rows.append(row)
        b.formulas("A%d:E%d" % (first, last), [[f.format(r=r) for f in row] for r, row in zip(range(first, last + 1), rows)])

        calc = []
        for r in range(first, last + 1):
            p = r - 1
            firstrow = r == first
            prev_bal = "RFS_OpeningBalance" if firstrow else "$M%d" % p
            prev_k = "RFS_PresentContribution" if firstrow else "$K%d" % p
            prev_z = "RFS_OpeningBalance" if firstrow else "$Z%d" % p
            prev_y = None if firstrow else "$Y%d" % p
            row = [
                '=IF($F{r}<>"";$F{r}-$B{r};IF(ISNA(MATCH($A{r};Exp_RevisedYears;0));0;'
                "-SUM(INDEX(Exp_Revised;0;MATCH($A{r};Exp_RevisedYears;0)))-$B{r}))",  # H
                "=$B{r}+$H{r}",  # I
                '=IF($A{r}<=PB_ContribFixedThrough;"";IF($A{r}-PB_ContribFixedThrough<=3;'
                "INDEX(PB_CatchUp;$A{r}-PB_ContribFixedThrough);PB_SubsequentIncrease))",  # J
                "=IF($A{r}<=PB_ContribFixedThrough;$D{r};%s*(1+$J{r}))" % prev_k,  # K
                "=%s*%s" % (prev_bal, rate),  # L
                '=IF($G{r}<>"";$G{r};%s+$I{r}+$L{r}+$K{r})' % prev_bal,  # M
                "=$M{r}-$E{r}",  # N
                "=(1+RFS_ShortInflation)^($A{r}-RFS_StudyYear)",  # O
                "=$M{r}-RFS_MinBalance",  # P
                "=$M{r}-RFS_MinBalance*$O{r}",  # Q
                "=$M{r}/$O{r}-RFS_MinBalance",  # R
                "=$E{r}-RFS_MinBalance",  # S
                "=$E{r}-RFS_MinBalance*$O{r}",  # T
                "=$E{r}/$O{r}-RFS_MinBalance",  # U
                "=$K{r}/12/RFS_Units",  # V
            ]
            checks = [
                "=%s*%s" % (prev_z, rate),  # X
                "=RFS_PresentContribution"
                if firstrow
                else "=%s*(1+IF($A{r}-RFS_StudyYear<=3;INDEX(RFS_CatchUpIncreases;$A{r}-RFS_StudyYear);RFS_SubsequentIncrease))" % prev_y,  # Y
                "=%s+$B{r}+$X{r}+$Y{r}" % prev_z,  # Z
                '=IF(ABS($Z{r}-$E{r})>1;"Differs from RFS";"")',  # AA
            ]
            calc.append(([f.format(r=r) for f in row], [f.format(r=r) for f in checks]))
        b.formulas("H%d:V%d" % (first, last), [c[0] for c in calc])
        b.formulas("X%d:AA%d" % (first, last), [c[1] for c in calc])

        # Inputs carried over from the simulator: the 2024 actual closing balance (H15) and the
        # 2025 actual expenditures (C16 = -1824525 - B16).
        b.value("G%d" % first, num(self.sim_cell("H", 15)))
        b.value("F%d" % (first + 1), -1824525.0)
        b.style("F%d:G%d" % (first, last), "Input")
        for row, c, text in self.sim_comments:
            if (row, c) == (15, colnum("H")):
                b.comment("G%d" % first, text)
            else:
                b.comment("A%d" % first, "[Comment from old simulator %s%d] %s" % (col(c), row, text))
        self.notes.append(
            "Projected_Balances: merges cashflow_spatac and cashflow_spatac_simulator. The 2025 "
            "actual expenditures (-1,824,525, hard-coded in the old C16) and the 2024 actual "
            "closing balance are now inputs; the unused 1st-year increase (old L9, 4.4%) is gone; "
            "the RFS balance is read from RFS_cashflow instead of pasted values."
        )

        b.fmt("B%d:I%d" % (first, last), MONEY_ZERO)
        b.fmt("J%d:J%d" % (first, last), PERCENT)
        b.fmt("K%d:N%d" % (first, last), MONEY_ZERO)
        b.fmt("O%d:O%d" % (first, last), "0.0000")
        b.fmt("P%d:U%d" % (first, last), MONEY_ZERO)
        b.fmt("V%d:V%d" % (first, last), "#,##0.00")
        b.fmt("X%d:Z%d" % (first, last), MONEY_ZERO)
        b.fmt("A%d:A%d" % (first, last), YEAR)
        b.width("A", 30)
        for k in range(2, 28):
            b.width(col(k), 13)
        b.width("A", 8)
        b.width("D", 13)
        b.width("G", 30)
        b.group_cols("X", "AA", collapse=True)

        span = lambda c: ref("Projected_Balances", "%s%d:%s%d" % (c, first, c, last))
        for n, c in (
            ("PB_Year", "A"), ("PB_RFSCost", "B"), ("PB_RFSInterest", "C"), ("PB_RFSContrib", "D"),
            ("PB_RFSBalance", "E"), ("PB_ActualCost", "F"), ("PB_ActualBalance", "G"),
            ("PB_Adjustment", "H"), ("PB_AdjustedCost", "I"), ("PB_Increase", "J"),
            ("PB_Contributions", "K"), ("PB_Interest", "L"), ("PB_Balance", "M"),
            ("PB_Variance", "N"), ("PB_Inflator", "O"), ("PB_Margin", "P"),
            ("PB_MarginInflated", "Q"), ("PB_MarginPV", "R"), ("PB_RFSMargin", "S"),
            ("PB_RFSMarginInflated", "T"), ("PB_RFSMarginPV", "U"), ("PB_Check", "AA"),
        ):
            define(doc, n, span(c))
        self.freeze(sheet, 1, 12)

    # -- order and protection ------------------------------------------------------------------

    def order_and_protect(self):
        sheets = self.doc.Sheets
        order = [
            "README", "Settings", "Expenditures", "Medium_Term_Cashflow",
            "Schedule_of_Investments", "Projected_Balances", "RFS_info", "RFS_fvalues",
            "RFS_cashflow",
        ]
        for i, name in enumerate(order):
            sheets.moveByName(name, i)
        for name in order[:6]:
            sheets.getByName(name).protect("")
        self.doc.CurrentController.setActiveSheet(sheets.getByName("README"))


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("source")
    ap.add_argument("output")
    ap.add_argument(
        "--fix",
        action="append",
        default=[],
        choices=["compounding", "closing"],
        help="apply a phase 4 bug fix (repeatable)",
    )
    args = ap.parse_args()
    with Session() as s:
        notes = Migration(s, args.source, args.output, set(args.fix)).run()
    print("Wrote %s" % args.output)
    for n in notes:
        print("- " + n)


if __name__ == "__main__":
    sys.exit(main())
