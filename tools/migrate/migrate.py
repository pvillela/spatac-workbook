"""Restructures the Spatac workbook as specified in docs/archive/design.md (phase 3).

Usage (LibreOffice's bundled Python carries the `uno` module):

    /opt/libreoffice26.2/program/python tools/migrate/migrate.py SOURCE.xlsx OUTPUT.xlsx

SOURCE is the pre-restructure workbook (Spatac-20261121.xlsx). It is opened read-only; nothing is
written to it. OUTPUT is a new workbook, built from scratch, with the vendor RFS tabs copied across
and every Spatac tab rebuilt with defined names, input styles and formulas that follow the design.
OUTPUT must not exist yet: the build never overwrites a file.

The migration reproduces today's results. Known bugs are preserved here and fixed separately
(phase 4, `--fix` options).
"""

import argparse
import os
import posixpath
import re
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

import uno  # noqa: F401  (registers the com.sun.star import hook)
from com.sun.star.table import CellRangeAddress
from com.sun.star.util import CellProtection

from reasons import reason_for
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
FIXED_COLOR = 0xF2F2F2

# The user's J9 border in the review copy: 0.26 mm (OuterLineWidth 26), dotted (LineStyle 1),
# grey 0x666666.
GRID_WIDTH = 26
GRID_STYLE = 1
GRID_COLOR = 0x666666
VERT_TOP = 1  # com.sun.star.table.CellVertJustify.TOP

# The Expenditures item sort key (design 13.5). The item number (text, e.g. 5.3.1a) gets four
# extra dots so that each FIND finds one; each part becomes its digits padded to four, then its
# letter suffix. Missing parts are empty, so 5.3 sorts before 5.3.1.
_PART = 'IF({s}="";"";"."&RIGHT("0000"&LEFT({s};LEN({s})-ISERROR(VALUE(RIGHT({s};1))));4)&MID({s};LEN({s})-ISERROR(VALUE(RIGHT({s};1)))+1;9))'
# LET names must not look like cell addresses (p1, s1 would be cells).
_WORDS = ["one", "two", "three", "four"]
SORT_KEY = (
    '=IF(TRIM($B{r}&"")="";"";LET(rest_zero;TRIM($B{r}&"")&"....";'
    + "".join(
        "pos_{w};FIND(\".\";rest_{p});part_{w};LEFT(rest_{p};pos_{w}-1);rest_{w};MID(rest_{p};pos_{w}+1;99);".format(
            w=w, p=(["zero"] + _WORDS)[i]
        )
        for i, w in enumerate(_WORDS)
    )
    + "MID(" + "&".join(_PART.format(s="part_" + w) for w in _WORDS) + ";2;99)))"
)

# Red negatives, as the user set on the Medium_Term_Cashflow variance rows.
MONEY_RED = "#,##0;[RED]-#,##0"

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

    def fix_heights(self, first, last, margin=1.1):
        """Explicit heights for wrapped rows (design 13.3). Excel does not re-fit a row saved as
        auto-height when it opens the file, so a wrapped cell would collapse to one line. Takes
        LibreOffice's fitted height, with a margin for Excel's slightly wider font metrics."""
        for r in range(first - 1, last):
            row = self.sheet.Rows.getByIndex(r)
            row.OptimalHeight = True
            row.Height = int(row.Height * margin)

    def group_cols(self, first, last, collapse=False):
        a = CellRangeAddress()
        a.Sheet = self.sheet.RangeAddress.Sheet
        a.StartColumn, a.EndColumn = colnum(first) - 1, colnum(last) - 1
        a.StartRow = a.EndRow = 0
        self.sheet.group(a, uno.Enum("com.sun.star.table.TableOrientation", "COLUMNS"))
        if collapse:
            self.sheet.hideDetail(a)

    # Comments wait until every formula has been rewritten (Migration.reparse): rewriting a
    # formula cell deletes its comment.
    pending_comments = []

    def comment(self, a, text):
        Builder.pending_comments.append((self.sheet, self.cell(a).CellAddress, text))


def define(doc, name, content):
    """A workbook-scoped name; `content` in Calc syntax, e.g. $Settings.$B$3."""
    names = doc.NamedRanges
    if names.hasByName(name):
        names.removeByName(name)
    origin = uno.createUnoStruct("com.sun.star.table.CellAddress")
    names.addNewByName(name, content, origin, 0)


def allow_autofilter(src, dst, sheet_names):
    """Copies the .xlsx `src` to `dst`, letting the AutoFilter be used on the named protected
    tabs (OOXML sheetProtection autoFilter="false")."""
    with zipfile.ZipFile(src) as z:
        book = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        rid = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        target = {r.get("Id"): r.get("Target") for r in rels}
        parts = set()
        for sh in book.find("{%s}sheets" % main):
            if sh.get("name") in sheet_names:
                parts.add("xl/" + target[sh.get(rid)].lstrip("/").replace("xl/", "", 1))
        if len(parts) != len(sheet_names):
            raise RuntimeError("tabs not found: %s" % sheet_names)
        with zipfile.ZipFile(dst, "x", zipfile.ZIP_DEFLATED) as out:
            for info in z.infolist():
                data = z.read(info.filename)
                if info.filename in parts:
                    text = data.decode("utf-8")
                    new, n = re.subn(r"<sheetProtection ([^>]*?)/>", r'<sheetProtection \1 autoFilter="false"/>', text)
                    if n != 1 or "autoFilter=" in re.search(r"<sheetProtection [^>]*?/>", text).group(0):
                        raise RuntimeError("unexpected sheet protection in %s" % info.filename)
                    data = new.encode("utf-8")
                out.writestr(info, data)


def ref(sheet, rng):
    """Absolute Calc reference: ref('Settings', 'B3') -> $Settings.$B$3."""
    def absolute(part):
        letters = "".join(ch for ch in part if ch.isalpha())
        digits = "".join(ch for ch in part if ch.isdigit())
        return "$%s$%s" % (letters, digits) if digits else "$%s" % letters
    parts = rng.split(":")
    quoted = "'%s'" % sheet if not sheet.replace("_", "").isalnum() else sheet
    return "$%s.%s" % (quoted, ":".join(absolute(p) for p in parts))


def grid_border():
    """The thin dotted grey border the user set on Expenditures J9 (design 13.2)."""
    line = uno.createUnoStruct("com.sun.star.table.BorderLine2")
    line.OuterLineWidth = GRID_WIDTH
    line.LineWidth = GRID_WIDTH
    line.LineStyle = GRID_STYLE
    line.Color = GRID_COLOR
    return line


def solid_border():
    """A thin solid black line, as on the user's year-block boxes and closing-balance rules."""
    line = uno.createUnoStruct("com.sun.star.table.BorderLine2")
    line.OuterLineWidth = GRID_WIDTH
    line.LineWidth = GRID_WIDTH
    return line


def make_styles(doc):
    styles = doc.StyleFamilies.getByName("CellStyles")
    # Top vertical alignment everywhere (design 13.2): the default style is Excel's "Normal".
    styles.getByName("Default").VertJustify = VERT_TOP

    def add(name, color=None, locked=True, bold=False, grid=False):
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
        if grid:
            for side in ("TopBorder", "BottomBorder", "LeftBorder", "RightBorder"):
                setattr(s, side, grid_border())
        return s

    add("Input", INPUT_COLOR, locked=False, grid=True)
    add("Calculated", CALC_COLOR, grid=True)
    # Amounts fixed for the year, such as the current year's Approved column (design 13.4).
    add("Fixed", FIXED_COLOR, grid=True)
    add("Header", HEADER_COLOR, bold=True)
    title = add("Title", bold=True)
    title.CharHeight = 14.0


# --------------------------------------------------------------------------------------------
# Reading the source


def read_block(sheet, rng):
    return [list(r) for r in sheet.getCellRangeByName(rng).getDataArray()]


def read_comments(path, sheet_name):
    """The comments of one tab of an .xlsx file, as (row, column, text), 1-based.

    Read from the file's XML because LibreOffice's API returns these comments with empty text
    (design 13.4). Excel's threaded comments are stored with a legacy stub ("[Threaded
    comment] ... Comment:") in front of the text; only the text after it is kept."""
    ns = {
        "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "p": "http://schemas.openxmlformats.org/package/2006/relationships",
    }
    with zipfile.ZipFile(path) as z:
        book = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = {r.get("Id"): r.get("Target") for r in rels.findall("p:Relationship", ns)}
        sheet_path = None
        for sh in book.find("m:sheets", ns):
            if sh.get("name") == sheet_name:
                sheet_path = "xl/" + target[sh.get("{%s}id" % ns["r"])].lstrip("/").replace("xl/", "", 1)
        if sheet_path is None:
            raise RuntimeError("no tab %s in %s" % (sheet_name, path))
        rels_path = posixpath.join(posixpath.dirname(sheet_path), "_rels", posixpath.basename(sheet_path) + ".rels")
        if rels_path not in z.namelist():
            return []
        comments_path = None
        for r in ET.fromstring(z.read(rels_path)).findall("p:Relationship", ns):
            if r.get("Type").endswith("/comments"):
                comments_path = posixpath.normpath(posixpath.join(posixpath.dirname(sheet_path), r.get("Target")))
        if comments_path is None:
            return []
        root = ET.fromstring(z.read(comments_path))
    out = []
    for c in root.iter("{%s}comment" % ns["m"]):
        text = "".join(t.text or "" for t in c.iter("{%s}t" % ns["m"]))
        if text.startswith("[Threaded comment]") and "Comment:" in text:
            text = text.split("Comment:", 1)[1]
        text = text.strip()
        m = re.match(r"([A-Z]+)(\d+)$", c.get("ref"))
        out.append((int(m.group(2)), colnum(m.group(1)), text))
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
        doc = self.s.new_calc(hidden=False)  # a hidden document loses its frozen panes
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
        for sheet, addr, text in Builder.pending_comments:
            sheet.Annotations.insertNew(addr, text)
        Builder.pending_comments.clear()
        self.order_and_protect()

        doc.calculateAll()
        # LibreOffice cannot set the "use AutoFilter" protection option through its API, so the
        # file is saved to a temporary path and written to the output once, with the option
        # set (design 13.5).
        with tempfile.TemporaryDirectory(prefix="spatac-migrate-") as tmp:
            saved = os.path.join(tmp, "out.xlsx")
            self.s.save_xlsx(doc, saved)
            allow_autofilter(saved, self.out_path, ["Expenditures"])
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
        self.fs_formulas = [
            list(r) for r in sh.getByName("fvalues_spatac").getCellRangeByName("A1:AP255").getFormulaArray()
        ]
        self.fs_comments = read_comments(self.src_path, "fvalues_spatac")
        self.mtc = read_block(sh.getByName("Medium_Term_Cashflow"), "A1:M69")
        self.inv = read_block(sh.getByName("Schedule_of_Investments"), "A1:T46")
        self.sim = read_block(sh.getByName("cashflow_spatac_simulator"), "A1:V65")
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
        b.value("A1", "CONFIGURATION SETTINGS").CellStyle = "Title"
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
        b.value("A9", "NOTE: The minimum target is inflated from its year at the RFS inflation rates.")
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
            ("Expenditures — every RFS element, with Spatac's approved amounts, actuals, estimates, board’s reasons for adjustments to the RFS, the RFS amounts by year, and the Revised amounts the plan uses.", None),
            ("Medium_Term_Cashflow — the Investment Plan: six years of cash, expenditures, maturities and planned new investments against the minimum target cash balances.", None),
            ("Schedule_of_Investments — investments held and planned, with each one's cash flows by year.", None),
            ("Projected_Balances — the reserve fund balance projected to the end of the RFS, with Spatac's adjustments and “what-if” contribution increases.", None),
            ("RFS_info, RFS_fvalues, RFS_cashflow — the vendor's RFS spreadsheet, pasted in unchanged. Helper columns to the right are not part of the paste.", None),
            ("Settings — the base year, the minimum target cash balance and the rate assumed for planned new investments.", None),
            ("", None),
            ("Cell colours", "Header"),
            ("Yellow — a manual input: user may type here.", "Input"),
            ("Green — calculated next to inputs: read only.", "Calculated"),
            ("Grey — fixed for the year, such as the current year's Approved amounts: changed only at the January rollover.", "Fixed"),
            ("Everything else is calculated or a label. Sheets are protected without a password to prevent accidents; Review > Unprotect Sheet when you need to.", None),
            ("", None),
            ("Expenditures tab: Estimates", "Header"),
            ("Mid-year and Late-year estimates are entered only for items whose estimate changed: type the new full amount, never a difference. Blank means no change; the Current estimate columns show the result for every item.", None),
            ("Spatac row — S marks an item Spatac added that is not in the RFS, such as a project the RFS does not foresee. These rows are exempt from the RFS check.", None),
            ("RFS check — flags a row whose item and GL account are not in RFS_info (\"Not in RFS\"), or whose item and element have no amounts in RFS_fvalues (\"No RFS amounts\"). Blank means the row matches the RFS. Check it after pasting a new RFS.", None),
            ("Item number, building element and GL account are fixed: unprotect the sheet to add or remove an item.", None),
            ("", None),
            ("Notice for Excel users", "Header"),
            ("Excel users: enter formulas normally (no Ctrl+Shift+Enter) and do not use formulas that spill; LibreOffice cannot keep them.", None),
        ]
        for i, (text, style) in enumerate(lines):
            c = b.value("A%d" % (i + 1), text)
            if style:
                c.CellStyle = style
        b.width("A", 120)
        b.wrap("A1:A%d" % len(lines))
        b.fix_heights(1, len(lines))

    # -- Expenditures --------------------------------------------------------------------------

    def build_expenditures(self):
        doc = self.doc
        doc.Sheets.insertNewByName("Expenditures", 2)
        sheet = doc.Sheets.getByName("Expenditures")
        b = Builder(doc, sheet)
        # Rows: 3 vendor totals, 4 totals, 5 variance vs RFS, 6 provenance notes (and the years
        # of the RFS and Revised blocks), 7 headers, then the items and the ▼ End row.
        # The row-block names run from the header row to the ▼ End row, so a row inserted just
        # below the header or just above ▼ End is inside them (design 13.5). The years sit in
        # row 6 rather than in the header row, so that no name takes in a number that is not
        # an amount.
        hdr, yr = 7, 6
        first = 8
        # The items: old rows 9 onward, through the last one in use. The old trailing blank
        # rows are dropped; new items are inserted above ▼ End.
        old_last = max(
            old for old in range(9, 253) if any(not is_blank(self.fs_cell(c, old)) for c in "ABCDE")
        )
        last = old_last - 1
        end = last + 1  # the ▼ End row
        sort_col = col(colnum("BZ") + 1)  # CA: the item sort key
        rfs0, rev0 = colnum("W"), colnum("AY")
        rfs_cols = [col(rfs0 + k) for k in range(RFS_YEARS)]
        rev_cols = [col(rev0 + k) for k in range(RFS_YEARS)]

        # Titles and group labels. The title is the user's B1 from the review copy; A1 is
        # left empty now that column A is visible.
        b.value("B1", "EXPENDITURES").CellStyle = "Title"
        b.formula("J2", '=Cfg_BaseYear&" — current year"')
        b.formula("O2", '=(Cfg_BaseYear+1)&" — next year"')
        b.formula("R2", '=(Cfg_BaseYear+2)&" — year after"')
        b.value("W2", "RFS amounts by year (from RFS_fvalues)")
        b.value("AY2", "Revised amounts by year (what the plan uses)")
        b.bold("A2:BZ2")

        # Row 3: vendor totals (a check), labelled next to the figures; row 4: totals; row 5:
        # variance vs RFS; row 6: provenance notes.
        b.value("V3", "Vendor total")
        b.cell("V3").HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", "RIGHT")
        b.value("C4", "Total")
        b.value("C5", "Variance vs RFS (more than the RFS is positive)")
        b.value("C6", "As of / source (edit when you update a column)")
        for k, c in enumerate(rfs_cols):
            b.formula(
                "%s3" % c,
                "=IFERROR(-INDEX(RFS_Fv_Totals;1;MATCH(%s$%d;RFS_Fv_Years;0));0)" % (c, yr),
            )
        # Totals only where they mean something: the change-only estimate columns (L, M, O, P,
        # R, S) hold amounts for changed items only, so they get none (design 13.4).
        for c in ["J", "K", "N", "Q", "T"] + rfs_cols + rev_cols:
            b.formula("%s4" % c, "=SUM(%s%d:%s%d)" % (c, hdr, c, end))
        # Variance: each Current estimate and Revised total against the RFS total of its year.
        rfs_total = "INDEX($%s$4:$%s$4;MATCH(%%s;Exp_RFSYears;0))" % (rfs_cols[0], rfs_cols[-1])
        for c, k in (("N", 0), ("Q", 1), ("T", 2)):
            b.formula("%s5" % c, "=%s4-%s" % (c, rfs_total % ("Cfg_BaseYear+%d" % k)))
        for c in rev_cols:
            b.formula("%s5" % c, "=%s4-%s" % (c, rfs_total % ("%s$%d" % (c, yr))))
        b.fmt("J3:BZ5", MONEY_ZERO)
        for c in ["N", "Q", "T"]:
            b.fmt("%s5" % c, MONEY_RED)
        b.fmt("%s5:%s5" % (rev_cols[0], rev_cols[-1]), MONEY_RED)
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
            b.value("%s6" % c, text)
            b.style("%s6" % c, "Input")
        b.wrap("A6:V6")

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
            "V": "Working notes",
        }
        for c, t in headers.items():
            b.value("%s%d" % (c, hdr), t)
        # The reasons explain the next year's approved changes only (design 13.4).
        b.formula("U%d" % hdr, '=(Cfg_BaseYear+1)&" Reason for change"')
        year_headers = {
            "J": "Approved",
            "K": "Actuals to date",
            "L": "Mid-year estimate",
            "M": "Late-year estimate",
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
            b.formula("%s%d" % (c, hdr), '=(Cfg_BaseYear+%d)&" %s"' % (offset[c], t))
        b.formula("W%d" % yr, "=Cfg_BaseYear")
        b.formula("AY%d" % yr, "=Cfg_BaseYear")
        for k in range(1, RFS_YEARS):
            b.formula("%s%d" % (rfs_cols[k], yr), "=%s%d+1" % (rfs_cols[k - 1], yr))
            b.formula("%s%d" % (rev_cols[k], yr), "=%s%d+1" % (rev_cols[k - 1], yr))
        for k in range(RFS_YEARS):
            b.formula("%s%d" % (rfs_cols[k], hdr), '="RFS "&%s$%d' % (rfs_cols[k], yr))
            b.formula("%s%d" % (rev_cols[k], hdr), '="Revised "&%s$%d' % (rev_cols[k], yr))
        b.value("%s%d" % (sort_col, hdr), "Item sort key")
        b.style("A%d:%s%d" % (hdr, sort_col, hdr), "Header")
        b.style("W%d:BZ%d" % (yr, yr), "Header")
        b.wrap("A%d:%s%d" % (hdr, sort_col, hdr))
        b.height(hdr, 42)

        # Data rows.
        identity, inputs, notes = [], [], []
        zero_rule = {"M": 0, "P": 0, "S": 0}
        spatac_rows = 0
        for r in range(first, last + 1):
            old = r + 1
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
        restored = self.restore_estimate_formulas(b, first, last)

        # Reasons for change, inferred from the old Working notes (reasons.py); a doubt becomes a
        # comment on the reason cell.
        written = doubts = 0
        for k, row in enumerate(identity):
            r = first + k
            label = trim(lo_text(row[2])) or trim(lo_text(row[4]))
            found = reason_for(trim(lo_text(row[1])), label)
            if found is None:
                continue
            reason, doubt = found
            b.value("U%d" % r, reason)
            written += 1
            if doubt:
                b.comment("U%d" % r, doubt)
                doubts += 1
        self.notes.append(
            "Expenditures: %d reasons for change inferred from the Working notes; %d carry a "
            "comment with a doubt to resolve." % (written, doubts)
        )
        self.notes.append(
            "Expenditures: %d rows marked as Spatac rows (no RFS item). Explicit zeros written "
            "where a blank used to mean zero: %d current-year Late-year, %d next-year Late-year, %d year-after "
            "Late-year estimates. %d estimate formulas carried over, rewritten for the new "
            "columns." % (spatac_rows, zero_rule["M"], zero_rule["P"], zero_rule["S"], restored)
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
                    "=IFERROR(INDEX(RFS_Fv_Values;MATCH($G{r};RFS_Fv_Key;0);MATCH({c}${h};RFS_Fv_Years;0));0)".format(r=r, c=c, h=yr)
                    for c in rfs_cols
                ]
                for r in range(first, last + 1)
            ],
        )
        b.formulas(
            "%s%d:%s%d" % (rev_cols[0], first, rev_cols[-1], last),
            [
                [
                    "=IF({c}${h}=Cfg_BaseYear;$N{r};IF({c}${h}=Cfg_BaseYear+1;$Q{r};IF({c}${h}=Cfg_BaseYear+2;$T{r};{f}{r})))".format(
                        r=r, c=c, f=f, h=yr
                    )
                    for c, f in zip(rev_cols, rfs_cols)
                ]
                for r in range(first, last + 1)
            ],
        )

        # Styles and formats. The identity columns A-F are protected (design 13.4): adding or
        # removing an item means unprotecting the sheet. Approved is fixed for the year.
        b.style("J%d:J%d" % (first, last), "Fixed")
        b.style("K%d:M%d" % (first, last), "Input")
        b.style("O%d:P%d" % (first, last), "Input")
        b.style("R%d:S%d" % (first, last), "Input")
        b.style("U%d:V%d" % (first, last), "Input")
        b.style("G%d:I%d" % (first, last), "Calculated")
        for c in "NQT":
            b.style("%s%d:%s%d" % (c, first, c, last), "Calculated")
        b.fmt("J%d:T%d" % (first, last), MONEY)
        b.fmt("%s%d:%s%d" % (rfs_cols[0], first, rev_cols[-1], last), MONEY)
        b.fmt("W%d:BZ%d" % (yr, yr), YEAR)
        widths = {"A": 6, "B": 8, "C": 34, "D": 10, "E": 24, "F": 6, "G": 20, "H": 16, "I": 8, "U": 30, "V": 40}
        for c in "JKLMNOPQRST":
            widths[c] = 11
        for c in rfs_cols + rev_cols:
            widths[c] = 10
        for c, w in widths.items():
            b.width(c, w)
        b.wrap("U%d:V%d" % (first, last))
        b.wrap("I%d:I%d" % (first, last))  # the narrower RFS check column

        # Boxes round the three year blocks' header areas, like the user's J2:N5 (review copy).
        for a, z in (("J", "N"), ("O", "Q"), ("R", "T")):
            box = b.cell("%s2:%s%d" % (a, z, hdr - 1))
            bd = uno.createUnoStruct("com.sun.star.table.TableBorder2")
            for side in ("TopLine", "BottomLine", "LeftLine", "RightLine"):
                setattr(bd, side, solid_border())
                setattr(bd, "Is%sValid" % side, True)
            box.TableBorder2 = bd

        # Groups (Q46, review copy): GL account name, Spatac row and helper keys in one
        # collapsed group; each year's inputs, and the two year blocks. Row # stays visible so
        # that sorting by it restores the original order (design 13.4).
        b.group_cols("E", "H", collapse=True)
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
            new_row = hdr if row <= 8 else row - 1
            if target is None:
                target = "V"
                text = "[Comment from old column %s] %s" % (letters, text)
            b.comment("%s%d" % (target, new_row), text)

        # Item sort key (design 13.5): the item number with each part padded, so that a text
        # sort puts 2.1 before 10.1 and 5.3.1a after 5.3.1. Up to four parts, each a number
        # of up to four digits with an optional letter suffix.
        b.formulas(
            "%s%d:%s%d" % (sort_col, first, sort_col, last),
            [[SORT_KEY.format(r=r)] for r in range(first, last + 1)],
        )
        b.style("%s%d:%s%d" % (sort_col, first, sort_col, last), "Calculated")
        b.width(sort_col, 14)
        b.group_cols(sort_col, sort_col, collapse=True)

        # The ▼ End row: inside every row-block name, no amounts and no sort keys, so any sort
        # leaves it at the bottom.
        b.value("C%d" % end, "▼ End of items: insert new rows above this line")
        b.fill("A%d:%s%d" % (end, sort_col, end), SEPARATOR_COLOR)
        b.bold("C%d" % end)

        # AutoFilter on the header row, usable while the tab is protected (allow_autofilter).
        udb = doc.getPropertyValue("UnnamedDatabaseRanges")
        udb.setByTable(b.cell("A%d:%s%d" % (hdr, sort_col, end)).RangeAddress)
        udb.getByTable(sheet.RangeAddress.Sheet).AutoFilter = True

        # Names: from the header row to the ▼ End row.
        def name(n, letters_from, letters_to=None):
            letters_to = letters_to or letters_from
            define(doc, n, ref("Expenditures", "%s%d:%s%d" % (letters_from, hdr, letters_to, end)))

        for n, c in (
            ("Exp_RowNumber", "A"), ("Exp_Item", "B"), ("Exp_Element", "C"), ("Exp_GL", "D"),
            ("Exp_GLName", "E"), ("Exp_SpatacRow", "F"), ("Exp_Key", "G"), ("Exp_Key2", "H"),
            ("Exp_Check", "I"), ("Exp_Approved", "J"), ("Exp_Actual", "K"), ("Exp_MidEst", "L"),
            ("Exp_LateEst", "M"), ("Exp_Current", "N"), ("Exp_MidEst1", "O"), ("Exp_LateEst1", "P"),
            ("Exp_Current1", "Q"), ("Exp_MidEst2", "R"), ("Exp_LateEst2", "S"), ("Exp_Current2", "T"),
            ("Exp_Reason", "U"), ("Exp_Notes", "V"), ("Exp_SortKey", sort_col),
        ):
            name(n, c)
        name("Exp_RFS", rfs_cols[0], rfs_cols[-1])
        name("Exp_Revised", rev_cols[0], rev_cols[-1])
        define(doc, "Exp_RFSYears", ref("Expenditures", "%s%d:%s%d" % (rfs_cols[0], yr, rfs_cols[-1], yr)))
        define(doc, "Exp_RevisedYears", ref("Expenditures", "%s%d:%s%d" % (rev_cols[0], yr, rev_cols[-1], yr)))

        # Frozen through column H and the header row, as in the review copy.
        self.freeze(sheet, colnum("H"), hdr)

    def restore_estimate_formulas(self, b, first, last):
        """Writes the old estimate formulas into their new cells, with their references
        rewritten: the old input columns I-L become J-M, the old pasted 2025 and 2026 RFS
        columns (G, H) become lookups of that item's RFS amount for that year, and the old
        inflation cell becomes RFS_ShortInflation. A reference to anything else stops the
        migration, rather than carrying over a formula that would point at the wrong cell."""
        newcol = {"I": "J", "J": "K", "K": "L", "L": "M", "N": "O", "O": "R", "P": "P", "Q": "S"}
        rfs_year = {"G": 2025, "H": 2026}
        inflation = re.compile(r"\$?'?cashflow_spatac'?\.\$?K\$?7")
        cell = re.compile(r"(?<![A-Za-z_.$])\$?([A-Z]{1,2})\$?(\d+)(?![\d(])")
        count = 0
        for r in range(first, last + 1):
            old = r + 1
            for letters, target in newcol.items():
                f = self.fs_formulas[old - 1][colnum(letters) - 1]
                if not f.startswith("="):
                    continue
                text = inflation.sub("RFS_ShortInflation", f)

                def rewrite(m):
                    c, row = m.group(1), int(m.group(2))
                    if row != old:
                        raise RuntimeError("fvalues_spatac!%s%d refers to another row: %s" % (letters, old, f))
                    if c in rfs_year:
                        return (
                            "IFERROR(INDEX(RFS_Fv_Values;MATCH($G%d;RFS_Fv_Key;0);MATCH(%d;RFS_Fv_Years;0));0)"
                            % (r, rfs_year[c])
                        )
                    if c in newcol:
                        return "%s%d" % (newcol[c], r)
                    raise RuntimeError("fvalues_spatac!%s%d refers to column %s: %s" % (letters, old, c, f))

                b.formula("%s%d" % (target, r), cell.sub(rewrite, text))
                count += 1
        return count

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
        b.cell("A1").CharHeight = 14.0
        b.formula("B3", "=Cfg_BaseYear")
        for k in range(1, HORIZON):
            b.formula("%s3" % yc[k], "=%s3+1" % yc[k - 1])
        b.style("A3:G3", "Header")
        b.fmt("B3:G3", YEAR)

        # Layout from the user's review copy (design 13.1): a blank row after the planned
        # maturities, the grid in rows 14-18 and the closing balance in row 20.
        # Phase 4 fix: the carry-over sits below the closing balance, then the adjustment.
        fixed = "closing" in self.fixes
        carry, adj, minimum = (22, 23, 24) if fixed else (19, 22, 24)
        newtot, grid0, closing = 13, 14, 20  # grid0..grid0+4: the 1- to 5-year rows
        labels = {
            5: self.mtc_cell("A", 5),
            6: self.mtc_cell("A", 6),
            8: "Planned contributions (RFS)",
            9: "Revised expenditures (from Expenditures)",
            10: "Held investments maturing + interest payments (from Schedule_of_Investments)",
            11: "Planned investments maturing (from the grid below)",
            newtot: "Planned new investments",
            grid0: "1-year maturity",
            grid0 + 1: "2-year maturity",
            grid0 + 2: "3-year maturity",
            grid0 + 3: "4-year maturity",
            grid0 + 4: "5-year maturity (balancing figure)",
            carry: "Additional cash carry-over",
            closing: "Closing cash balance",
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
            b.formula("%s5" % yc[k], "=%s%d" % (yc[k - 1], closing))
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
                    "=SUMPRODUCT($B${g}:{p}{h}*((COLUMN($B${g}:{p}{h})-COLUMN($B${g}))+(ROW($B${g}:{p}{h})-ROW($B${g})+1)"
                    "=(COLUMN({c}{g})-COLUMN($B${g})))*(1+Cfg_NewInvRate)^(ROW($B${g}:{p}{h})-ROW($B${g})+1))".format(
                        p=prev, c=c, g=grid0, h=grid0 + 4
                    ),
                )
            b.formula("%s%d" % (c, newtot), "=-SUM(%s%d:%s%d)" % (c, grid0, c, grid0 + 4))
            for k in range(4):  # source rows 13-16
                b.value("%s%d" % (c, grid0 + k), self.mtc_cell(c, 13 + k))
                b.style("%s%d" % (c, grid0 + k), "Input")
            b.formula(
                "%s%d" % (c, grid0 + 4),
                "=SUM({c}4:{c}11)-SUM({c}{g}:{c}{h})-{c}{k}-{c}{m}-{c}{a}".format(
                    c=c, k=carry, m=minimum, a=adj, g=grid0, h=grid0 + 3
                ),
            )
            b.value("%s%d" % (c, carry), self.mtc_cell(c, 19))
            b.style("%s%d" % (c, carry), "Input")
            # Phase 4 fix: the balancing figure already leaves the carry-over and the adjustment
            # in cash, so the closing balance adds nothing more. The old formula added the
            # adjustment a second time.
            if fixed:
                b.formula("%s%d" % (c, closing), "=SUM({c}5:{c}{n})".format(c=c, n=newtot))
            else:
                b.formula("%s%d" % (c, closing), "=SUM({c}5:{c}{n})+{c}{a}".format(c=c, n=newtot, a=adj))
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
        for r in (30, 34):  # the variances: red negatives (review copy)
            b.fmt("B%d:G%d" % (r, r), MONEY_RED)
        for r in (closing, minimum):
            b.bold("A%d:G%d" % (r, r))
        # Rules from the review copy: under the planned new investments total, and above and
        # below the closing balance.
        rule = lambda rng, side: setattr(b.cell(rng), side, solid_border())
        rule("B%d:G%d" % (newtot, newtot), "BottomBorder")
        rule("B%d:G%d" % (closing, closing), "TopBorder")
        rule("B%d:G%d" % (closing, closing), "BottomBorder")
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
            ("MTC_NewInvTotal", "B%d:G%d" % (newtot, newtot)),
            ("MTC_NewInv", "B%d:G%d" % (grid0, grid0 + 4)),
            ("MTC_CarryOver", "B%d:G%d" % (carry, carry)),
            ("MTC_ClosingCash", "B%d:G%d" % (closing, closing)),
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
        # Block boundaries (design 13.5): the "Held" label row above the first held investment
        # and the ▼ End row below the last are inside the Held names, so a row inserted just
        # below the label or just above ▼ End is inside them.
        held_top, held_end = held0 - 1, held1 + 1  # 6, 47
        plan_hdr = held1 + 2  # 48
        plan0, plan1 = plan_hdr + 1, plan_hdr + PLANNED_ROWS  # 49..78
        total = plan1 + 2  # 80

        b.value("A1", self.inv_cell("A", 1)).CellStyle = "Input"
        b.bold("A1")
        b.cell("A1").CharHeight = 14.0  # review copy
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
        b.value("A%d" % held_end, "▼ End of held investments: insert new rows above this line")
        b.fill("A%d:X%d" % (held_end, held_end), SEPARATOR_COLOR)
        b.bold("A%d" % held_end)
        b.value("A%d" % total, "Total cash flows")
        for c in yc:
            b.formula("%s%d" % (c, total), "=SUM(%s%d:%s%d)" % (c, held_top, c, plan1))
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
            define(doc, n, span("%s%d" % (c, held_top), "%s%d" % (c, plan1)))
        define(doc, "Inv_Title", span("A1", "A1"))
        define(doc, "Inv_Held", span("A%d" % held_top, "X%d" % held_end))
        define(doc, "Inv_Planned", span("A%d" % plan0, "X%d" % plan1))
        define(doc, "Inv_Years", span("N5", "X5"))
        define(doc, "Inv_YearFlows", span("N%d" % held_top, "X%d" % plan1))
        define(doc, "Inv_HeldFlows", span("N%d" % held_top, "X%d" % held_end))
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
        """The reserve fund balance year by year (design 13.6): the RFS baseline, read from
        RFS_cashflow, and a what-if projection beside it. With the what-if's three choices set to
        RFS it recomputes the vendor's balances, which validates RFS_cashflow."""
        doc = self.doc
        doc.Sheets.insertNewByName("Projected_Balances", 5)
        sheet = doc.Sheets.getByName("Projected_Balances")
        b = Builder(doc, sheet)
        hdr = 14
        first, n_years = hdr + 1, 50
        last = first + n_years - 1  # 2024..2073
        last30 = first + 29  # the first 30 years of the RFS

        b.value("A1", "PROJECTED RESERVE FUND BALANCES").CellStyle = "Title"
        b.value("A2", "The RFS's own projection (columns B-H, from RFS_cashflow) and a what-if projection (columns J-S).")
        b.value(
            "A3",
            "To check the vendor's arithmetic, set the three what-if choices to RFS: the difference "
            "from the RFS balance (column O) is then zero in every year.",
        )

        # What-if inputs.
        inputs = [
            # (row, label, value, format, name, choices)
            (4, "What-if starts from", "Current year", None, "PB_Start", ("Current year", "RFS")),
            (5, "Opening balance of the current year (Current year start only)",
             num(self.sim_cell("H", 16)), MONEY_ZERO, "PB_OpeningBalance", None),
            (6, "Expenditures", "Revised", None, "PB_ExpBasis", ("Revised", "RFS")),
            (7, "Contributions", "Increases", None, "PB_ContribBasis", ("Increases", "RFS")),
            (8, "Contributions follow the RFS schedule through", 2027.0, YEAR, "PB_ContribFixedThrough", None),
            (9, "Contribution increase, 1st year after", num(self.sim_cell("Q", 9)), PERCENT, None, None),
            (10, "Contribution increase, 2nd year after", num(self.sim_cell("Q", 10)), PERCENT, None, None),
            (11, "Contribution increase, 3rd year after", num(self.sim_cell("Q", 11)), PERCENT, None, None),
            (12, "Subsequent annual increase", "=RFS_SubsequentIncrease", PERCENT, "PB_SubsequentIncrease", None),
        ]
        for r, label, v, fmt, name, choices in inputs:
            b.value("A%d" % r, label)
            if isinstance(v, str) and v.startswith("="):
                b.formula("B%d" % r, v)
            else:
                b.value("B%d" % r, v)
            b.style("B%d" % r, "Input")
            if fmt:
                b.fmt("B%d" % r, fmt)
            if name:
                define(doc, name, ref("Projected_Balances", "B%d" % r))
            if choices:
                cell = b.cell("B%d" % r)
                val = cell.Validation
                val.Type = uno.Enum("com.sun.star.sheet.ValidationType", "LIST")
                val.ShowList = 1
                val.ShowErrorMessage = True
                val.ErrorMessage = "Choose %s." % " or ".join(choices)
                val.setFormula1(";".join('"%s"' % c for c in choices))
                cell.Validation = val
        define(doc, "PB_CatchUp", ref("Projected_Balances", "B9:B11"))
        b.comment(
            "B5",
            "Seeded with the old simulator's projected closing balance for 2025, from the 2024 "
            "audited closing balance and the 2025 actual expenditures. Replace it with the "
            "figure from the financial statements: cash, less payables and accrued liabilities, "
            "plus investment principal and accrued interest.",
        )
        b.value("C4", "starts in")
        b.formula("D4", '=IF(PB_Start="RFS";RFS_StudyYear;Cfg_BaseYear)')
        b.fmt("D4", YEAR)
        define(doc, "PB_StartYear", ref("Projected_Balances", "D4"))

        b.value("D8", "Historical inflation rate 2021–2024")
        b.value("E8", num(self.sim_cell("T", 4)))
        b.value("D9", "Unlabelled value carried over from the old simulator (Q7)")
        b.value("E9", num(self.sim_cell("Q", 7)))
        b.fmt("E8:E9", PERCENT)

        # Lowest margins over the first 30 years of the RFS.
        b.value("G4", "Lowest over the first 30 years")
        b.value("H4", "Amount")
        b.value("I4", "Year")
        lows = (
            ("What-if: margin over minimum", "P"),
            ("What-if: margin over inflated minimum", "Q"),
            ("What-if: margin over minimum (PV)", "R"),
            ("RFS: margin over minimum", "F"),
            ("RFS: margin over inflated minimum", "G"),
            ("RFS: margin over minimum (PV)", "H"),
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

        # The year table.
        b.value("B%d" % (hdr - 1), "RFS (from RFS_cashflow; read only)")
        b.value("J%d" % (hdr - 1), "What-if")
        b.bold("A%d:S%d" % (hdr - 1, hdr - 1))
        heads = {
            "A": "Year",
            "B": "RFS expected costs", "C": "RFS interest", "D": "RFS contributions",
            "E": "RFS balance", "F": "RFS margin over minimum",
            "G": "RFS margin over inflated minimum", "H": "RFS margin over minimum (PV)",
            "J": "Expenditures", "K": "Contribution increase", "L": "Contributions",
            "M": "Interest", "N": "Closing balance", "O": "Difference from RFS balance",
            "P": "Margin over minimum", "Q": "Margin over inflated minimum",
            "R": "Margin over minimum (PV)", "S": "Contribution per unit per month",
            "U": "Inflator",
        }
        for c, h in heads.items():
            b.value("%s%d" % (c, hdr), h)
        b.style("A%d:U%d" % (hdr, hdr), "Header")
        b.wrap("A%d:U%d" % (hdr, hdr))
        b.height(hdr, 42)

        rate = "IF($A{r}<=RFS_StudyYear+3;RFS_ShortInterest;RFS_LongInterest)"
        rows = []
        for r in range(first, last + 1):
            p = r - 1
            firstrow = r == first
            if firstrow:
                opening = 'IF(PB_Start="RFS";RFS_OpeningBalance;PB_OpeningBalance)'
                prev_contrib = "RFS_PresentContribution"
            else:
                opening = 'IF($A{r}=PB_StartYear;IF(PB_Start="RFS";RFS_OpeningBalance;PB_OpeningBalance);$N%d)' % p
                prev_contrib = 'IF($L%d="";$D%d;$L%d)' % (p, p, p)
            active = "$A{r}>=PB_StartYear"
            row = [
                "=RFS_StudyYear" if firstrow else "=$A%d+1" % p,  # A
                "=IFERROR(INDEX(RFS_Cf_Cost;MATCH($A{r};RFS_Cf_Year;0));0)",  # B
                "=IFERROR(INDEX(RFS_Cf_Interest;MATCH($A{r};RFS_Cf_Year;0));0)",  # C
                "=IFERROR(INDEX(RFS_Cf_Contrib;MATCH($A{r};RFS_Cf_Year;0));0)",  # D
                "=IFERROR(INDEX(RFS_Cf_Balance;MATCH($A{r};RFS_Cf_Year;0));0)",  # E
                "=$E{r}-RFS_MinBalance",  # F
                "=$E{r}-RFS_MinBalance*$U{r}",  # G
                "=$E{r}/$U{r}-RFS_MinBalance",  # H
                "",  # I: spacer
                # J: the RFS cost, or the Revised amounts for the years the Revised block covers.
                '=IF(NOT(%s);"";IF(OR(PB_ExpBasis="RFS";ISNA(MATCH($A{r};Exp_RevisedYears;0)));$B{r};'
                "-SUM(INDEX(Exp_Revised;0;MATCH($A{r};Exp_RevisedYears;0)))))" % active,
                # K: the increase, after the years that follow the RFS schedule.
                '=IF(OR(NOT(%s);PB_ContribBasis="RFS";$A{r}<=PB_ContribFixedThrough);"";'
                "IF($A{r}-PB_ContribFixedThrough<=3;INDEX(PB_CatchUp;$A{r}-PB_ContribFixedThrough);"
                "PB_SubsequentIncrease))" % active,
                '=IF(NOT(%s);"";IF($K{r}="";$D{r};%s*(1+$K{r})))' % (active, prev_contrib),  # L
                '=IF(NOT(%s);"";%s*%s)' % (active, opening, rate),  # M
                '=IF(NOT(%s);"";%s+$J{r}+$M{r}+$L{r})' % (active, opening),  # N
                '=IF($N{r}="";"";$N{r}-$E{r})',  # O
                '=IF($N{r}="";"";$N{r}-RFS_MinBalance)',  # P
                '=IF($N{r}="";"";$N{r}-RFS_MinBalance*$U{r})',  # Q
                '=IF($N{r}="";"";$N{r}/$U{r}-RFS_MinBalance)',  # R
                '=IF($L{r}="";"";$L{r}/12/RFS_Units)',  # S
                "",  # T: spacer
                "=(1+RFS_ShortInflation)^($A{r}-RFS_StudyYear)",  # U
            ]
            rows.append([f.format(r=r) if f else "" for f in row])
        b.formulas("A%d:U%d" % (first, last), rows)

        b.style("J%d:S%d" % (first, last), "Calculated")
        b.fmt("A%d:A%d" % (first, last), YEAR)
        b.fmt("B%d:H%d" % (first, last), MONEY_ZERO)
        b.fmt("J%d:J%d" % (first, last), MONEY_ZERO)
        b.fmt("K%d:K%d" % (first, last), PERCENT)
        b.fmt("L%d:R%d" % (first, last), MONEY_ZERO)
        b.fmt("O%d:O%d" % (first, last), MONEY_RED)
        b.fmt("S%d:S%d" % (first, last), "#,##0.00")
        b.fmt("U%d:U%d" % (first, last), "0.0000")
        b.width("A", 52)
        for k in range(2, 22):
            b.width(col(k), 13)
        b.width("I", 2)
        b.width("T", 2)
        b.width("G", 30)
        b.group_cols("U", "U", collapse=True)
        self.notes.append(
            "Projected_Balances: rebuilt as the RFS baseline and a what-if projection (design "
            "13.6). The typed 2024 closing balance and 2025 actual expenditures are dropped; the "
            "old simulator's 2025 closing projection seeds the current-year opening balance. The "
            "hidden check columns are replaced by the what-if's difference from the RFS balance."
        )

        span = lambda c: ref("Projected_Balances", "%s%d:%s%d" % (c, first, c, last))
        for n, c in (
            ("PB_Year", "A"), ("PB_RFSCost", "B"), ("PB_RFSInterest", "C"), ("PB_RFSContrib", "D"),
            ("PB_RFSBalance", "E"), ("PB_RFSMargin", "F"), ("PB_RFSMarginInflated", "G"),
            ("PB_RFSMarginPV", "H"), ("PB_Cost", "J"), ("PB_Increase", "K"),
            ("PB_Contributions", "L"), ("PB_Interest", "M"), ("PB_Balance", "N"),
            ("PB_Variance", "O"), ("PB_Margin", "P"), ("PB_MarginInflated", "Q"),
            ("PB_MarginPV", "R"), ("PB_PerUnitMonth", "S"), ("PB_Inflator", "U"),
        ):
            define(doc, n, span(c))
        self.freeze(sheet, 1, hdr)

    # -- order and protection ------------------------------------------------------------------

    def order_and_protect(self):
        sheets = self.doc.Sheets
        # Settings goes last (design 13.2).
        order = [
            "README", "Expenditures", "Medium_Term_Cashflow", "Schedule_of_Investments",
            "Projected_Balances", "RFS_info", "RFS_fvalues", "RFS_cashflow", "Settings",
        ]
        for i, name in enumerate(order):
            sheets.moveByName(name, i)
        # Top alignment on the vendor tabs too (design 13.2): their cells carry the vendor's
        # own direct formatting, which the default style does not override.
        for name in ("RFS_info", "RFS_fvalues", "RFS_cashflow"):
            sh = sheets.getByName(name)
            cur = sh.createCursor()
            cur.gotoEndOfUsedArea(False)
            a = cur.RangeAddress
            sh.getCellRangeByPosition(0, 0, a.EndColumn, a.EndRow).VertJustify = VERT_TOP
        # The Spatac tabs are protected, without a password, as before (design 3.5, 13.2).
        for name in ("README", "Settings", "Expenditures", "Medium_Term_Cashflow",
                     "Schedule_of_Investments", "Projected_Balances"):
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
