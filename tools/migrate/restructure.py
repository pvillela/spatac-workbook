"""Restructures the Spatac workbook around an Adjustments tab (docs/archive/design.md section 14).

Usage (LibreOffice's bundled Python carries the `uno` module):

    /opt/libreoffice26.2/program/python tools/migrate/restructure.py \
        SOURCE.xlsx MOCKUP.xlsx OUTPUT.xlsx

SOURCE is the current build, _nogit/Spatac-20261123c.xlsx, which matches the user's golden source
_nogit/Spatac-L-20261123.xlsx. MOCKUP is the user's mock-up, _nogit/Spatac-20261123c-restructured.xlsx,
from which the user's adjustments (its 2026 block) are read. Both are opened read-only. OUTPUT
must not exist yet: the script never overwrites a file.

Unlike migrate.py, which builds every tab from scratch, this script edits SOURCE in place: what
section 14 does not change keeps its formats, comments, groups and names.
"""

import argparse
import os
import re
import sys
import tempfile
import zipfile

import uno  # noqa: F401  (registers the com.sun.star import hook)
from com.sun.star.table import CellRangeAddress
from com.sun.star.util import CellProtection

from migrate import (
    MONEY_RED,
    HEADER_COLOR,
    MONEY_ZERO,
    PERCENT,
    SEPARATOR_COLOR,
    YEAR,
    Builder,
    allow_autofilter,
    col,
    colnum,
    define,
    ref,
    solid_border,
)
from uno_session import Session

RFS_SPAN = 30  # the RFS planning span: study year .. study year + 29
SHORT_INFLATION_YEARS = 3  # "Inflation 0-3" in RFS_fvalues W1 (design 14.5)
SHORT_INTEREST_YEARS = 3  # "Short Term (3 yr.) Interest" in RFS_cashflow F7 (design 15.1)
BEYOND_COLOR = 0xEDEDED  # Projected_Balances: a year past the RFS's 30-year horizon
MONEY_CENTS_RED = "#,##0.00;[RED]-#,##0.00"  # MONEY_RED with cents (design 15.4)
CHANGED_COLOR = 0xFCE4D6  # an adjusted amount that differs from the RFS

# Expenditures after the restructure (design 14.6).
# The columns are built without the spacer K, which is inserted last (design 14.12). The final
# layout: J label, K spacer, L-P current year, Q notes, R spacer, S next year, T spacer, U onwards
# the adjusted RFS.
LABEL = "J"
ADJ_FIRST = colnum("T")  # 1-based, before the spacer

# Adjustments blocks (design 14.4): title, blank, header, input rows, closing row, two spacers.
BLOCKS = 4
BLOCK_ROWS = 10
FIRST_TITLE = 4
BLOCK_HEIGHT = 3 + BLOCK_ROWS + 1 + 2

# The user's adjustments, from the mock-up's 2026 block (A21:E26), restated in from-year
# dollars, or to-year dollars when Year from is "none" (design 14.4). Row 5's amount is the user's
# "difference between 600,000 and the above amount" as a formula; {prev} is the row above.
USER_AMOUNTS = [229500, 50000, 416160, 524979.5976, "=600000-G{prev}", 42448.32]

# Notes written by the golden source name its columns by letter: I Approved, J Actuals, K Mid-year,
# L Late-year. They now name them by their headings, which do not change when columns move
# (design 14.12).
# The headings are quoted, as the user asked.
COLUMN_NAMES = {
    "I": '"Approved"',
    "J": '"Actuals to date"',
    "K": '"Mid-year estimate"',
    "L": '"Late-year estimate"',
}
NAMED = "|".join(re.escape(n) for n in COLUMN_NAMES.values())

# Two notes that were already wrong in the golden source: "K = 0: not in YE projection" means the
# Late-year estimate, which is 0 on both rows (golden L, now N), not the Mid-year one. Corrected
# at the user's request, after the letters are translated.
NOTE_FIXES = {
    "11.5": ('"Mid-year estimate" = 0: not in YE projection', '"Late-year estimate" = 0: not in YE projection'),
    "18.11": ('"Mid-year estimate" = 0: not in YE projection', '"Late-year estimate" = 0: not in YE projection'),
}

COMMENT_39 = (
    "2028: the golden source (Spatac-L-20261123.xlsx) had 238,772, the 229,500 moved from 2025 "
    "inflated for two years. Three years' inflation gives 243,547."
)


def translate_note(text):
    """Replaces the column letters a working note names by the columns' headings, all at once,
    quoted, with spaces round the = and + next to them: K=J+pending reads
    "Mid-year estimate" = "Actuals to date" + pending."""
    text = re.sub(
        r"(?<![A-Za-z0-9’'.-])([IJKL])(?![A-Za-z0-9’'-])",
        lambda m: COLUMN_NAMES[m.group(1)],
        text,
    )
    text = re.sub(r"(%s) ?([=+]) ?" % NAMED, r"\1 \2 ", text)
    return re.sub(r" ?([=+]) ?(%s)" % NAMED, r" \1 \2", text)


def label_of(element, gl, gl_name, spatac):
    """The adjustment label of an Expenditures row (design 14.6), as its formula computes it."""
    gl = " ".join(str(gl).split())
    if not gl:
        return ""
    element = " ".join(str(element).split())
    core = "%s - %s" % (gl, element if element else " ".join(str(gl_name).split()))
    return core + " (Spatac)" if str(spatac).strip() else core


def keep_autofilter(saved, out_path):
    """Writes `saved` to `out_path` once. The AutoFilter stays usable on the protected
    Expenditures tab: LibreOffice keeps the option on a round trip, so it is only added when the
    save lost it (design 13.5)."""
    with zipfile.ZipFile(saved) as z:
        for info in z.infolist():
            if info.filename.startswith("xl/worksheets/sheet"):
                text = z.read(info.filename).decode("utf-8")
                if 'autoFilter="false"' in text:
                    with open(saved, "rb") as src, open(out_path, "xb") as dst:
                        dst.write(src.read())
                    return
    allow_autofilter(saved, out_path, ["Expenditures"])


def protection(locked):
    p = CellProtection()
    p.IsLocked = locked
    p.IsFormulaHidden = False
    p.IsHidden = False
    p.IsPrintHidden = False
    return p


def box(cell_range):
    bd = uno.createUnoStruct("com.sun.star.table.TableBorder2")
    for side in ("TopLine", "BottomLine", "LeftLine", "RightLine"):
        setattr(bd, side, solid_border())
        setattr(bd, "Is%sValid" % side, True)
    cell_range.TableBorder2 = bd


def remove_comment(sheet, address):
    target = sheet.getCellRangeByName(address).CellAddress
    notes = sheet.Annotations
    for i in range(notes.Count):
        a = notes.getByIndex(i).Position
        if a.Row == target.Row and a.Column == target.Column:
            notes.removeByIndex(i)
            return True
    return False


class Restructure:
    def __init__(self, session, src_path, mockup_path, out_path):
        self.s = session
        self.src_path = src_path
        self.mockup_path = mockup_path
        self.out_path = out_path
        self.report = []

    def run(self):
        if os.path.exists(self.out_path):
            raise SystemExit("%s exists: the script never overwrites a file" % self.out_path)
        self.user_rows = self.read_mockup()
        doc = self.s.open_copy(self.src_path, hidden=False)  # hidden documents lose frozen panes
        self.doc = doc
        doc.calculateAll()
        for sheet in doc.Sheets:
            if sheet.isProtected():
                sheet.unprotect("")
        self.add_styles()
        self.exp = doc.Sheets.getByName("Expenditures")
        self.locate_expenditures()

        self.settings()
        self.adjustments_tab()  # its names come first: Expenditures reads them
        self.expenditures()
        self.adjustments_logic()  # reads the Expenditures names
        self.user_adjustments()
        self.mtc()
        self.projected()
        self.readme()
        self.instruction_comments()
        self.no_hidden_zeros()
        self.order_and_protect()

        doc.calculateAll()
        with tempfile.TemporaryDirectory(prefix="spatac-restructure-") as tmp:
            saved = os.path.join(tmp, "out.xlsx")
            self.s.save_xlsx(doc, saved)
            keep_autofilter(saved, self.out_path)
        doc.close(True)
        return self.report

    # -- inputs -------------------------------------------------------------------------------

    def read_mockup(self):
        doc = self.s.open_copy(self.mockup_path)
        sheet = doc.Sheets.getByName("Adjustments")
        rows = [list(r) for r in sheet.getCellRangeByName("A21:E26").getDataArray()]
        doc.close(True)
        if len(rows) != len(USER_AMOUNTS) or any(not str(r[0]).strip() for r in rows):
            raise RuntimeError("the mock-up's 2026 block is not the six rows expected")
        return rows

    def add_styles(self):
        styles = self.doc.StyleFamilies.getByName("CellStyles")
        if not styles.hasByName("Changed"):
            s = self.doc.createInstance("com.sun.star.style.CellStyle")
            styles.insertByName("Changed", s)
            s.ParentStyle = "Default"
            s.CellBackColor = CHANGED_COLOR
            s.IsCellBackgroundTransparent = False

    def locate_expenditures(self):
        """Header row 7; items from 8 to the row before ▼ End."""
        data = self.exp.getCellRangeByName("C1:C400").getDataArray()
        ends = [i + 1 for i, r in enumerate(data) if str(r[0]).startswith("▼ End")]
        if len(ends) != 1:
            raise RuntimeError("Expenditures: no single ▼ End row")
        self.hdr, self.first, self.end = 7, 8, ends[0]
        self.last = self.end - 1

    # -- Settings -----------------------------------------------------------------------------

    def settings(self):
        sheet = self.doc.Sheets.getByName("Settings")
        b = Builder(self.doc, sheet)
        if b.cell("A8").getString() or b.cell("B8").getString():
            raise RuntimeError("Settings row 8 is not free")
        b.value("A8", "Years of short-term inflation in the RFS (the long-term rate applies after them)")
        b.value("B8", SHORT_INFLATION_YEARS)
        b.style("B8", "Input")
        b.fmt("B8", "0")
        b.cell("B8").HoriJustify = b.cell("B7").HoriJustify
        define(self.doc, "Cfg_ShortInflationYears", ref("Settings", "B8"))
        # A separate count for interest (design 15.1); the note below moves down a row.
        sheet.Rows.insertByIndex(8, 1)
        b.value("A9", "Years of short-term interest in the RFS (the long-term rate applies after them)")
        b.value("B9", SHORT_INTEREST_YEARS)
        b.style("B9", "Input")
        b.fmt("B9", "0")
        b.cell("B9").HoriJustify = b.cell("B8").HoriJustify
        b.cell("A9").CellStyle = b.cell("A8").CellStyle
        define(self.doc, "Cfg_ShortInterestYears", ref("Settings", "B9"))

    # -- Adjustments --------------------------------------------------------------------------

    def block_rows(self, k):
        title = FIRST_TITLE + k * BLOCK_HEIGHT
        header = title + 2
        return title, header, header + 1, header + BLOCK_ROWS, header + BLOCK_ROWS + 1

    def adjustments_tab(self):
        doc = self.doc
        doc.Sheets.insertNewByName("Adjustments", 2)
        sheet = doc.Sheets.getByName("Adjustments")
        self.adj = sheet
        b = Builder(doc, sheet)
        top = self.block_rows(0)[1]
        bottom = self.block_rows(BLOCKS - 1)[4]
        self.adj_top, self.adj_bottom = top, bottom

        # Names over every block, from the first header row to the last closing row, so a row
        # inserted inside any block is inside them.
        for n, c in (
            ("Adj_Item", "A"), ("Adj_From", "B"), ("Adj_To", "C"), ("Adj_Amount", "D"),
            ("Adj_Reason", "E"), ("Adj_FromAmount", "F"), ("Adj_ToAmount", "G"), ("Adj_Check", "H"),
        ):
            define(doc, n, ref("Adjustments", "%s%d:%s%d" % (c, top, c, bottom)))
        # The choices for Year from and Year to: "none" and the RFS's 30 years, in hidden J.
        b.value("J%d" % top, "none")
        b.formulas(
            "J%d:J%d" % (top + 1, top + RFS_SPAN),
            [["=RFS_StudyYear+%d" % k] for k in range(RFS_SPAN)],
        )
        define(doc, "Adj_YearChoices", ref("Adjustments", "J%d:J%d" % (top, top + RFS_SPAN)))

        b.value("A1", "ADJUSTMENTS TO PLANNED EXPENDITURES").CellStyle = "Title"  # the mock-up's title
        b.formula(
            "A2",
            "=LET(\n"
            "  problems; SUMPRODUCT((LEN(Adj_Check)>0)*(Adj_Check<>\"Check\")*(Adj_Check<>\"OK\"));\n"
            "  IF(problems=0; \"\"; \"Check: \"&problems&\" row(s) need attention; see column H \"&\n"
            "    \"and The Adjustments tab in the documentation.\"))",
        )
        b.cell("A2").CharColor = 0xC00000
        b.bold("A2")

        headers = [
            "Item (GL account - element)",
            "Year from",
            "Year to",
            "Amount (from-year $; to-year $ when Year from is none)",
            "Reason for change",
            "From-year amount",
            "To-year amount",
            "Check",
        ]
        for k in range(BLOCKS):
            title, header, r1, r2, closing = self.block_rows(k)
            b.formula("A%d" % title, '="Approved in "&(RFS_StudyYear+%d)' % k)
            b.bold("A%d" % title)
            b.cell("A%d" % title).CharHeight = 12.0
            b.values("A%d:H%d" % (header, header), [headers])
            b.style("A%d:H%d" % (header, header), "Header")
            b.wrap("A%d:H%d" % (header, header))
            b.height(header, 30)
            b.style("A%d:E%d" % (r1, r2), "Input")
            b.style("F%d:H%d" % (r1, r2), "Calculated")
            b.fill("A%d:H%d" % (closing, closing), SEPARATOR_COLOR)
            b.fmt("B%d:C%d" % (r1, r2), YEAR)
            b.fmt("D%d:G%d" % (r1, r2), MONEY_ZERO)
            b.wrap("E%d:E%d" % (r1, r2))
            b.wrap("H%d:H%d" % (r1, r2))
            # Right-justified, so that the OKs keep clear of the To-year amounts.
            b.cell("H%d:H%d" % (r1, r2)).HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", "RIGHT")
        for c, w in {"A": 46, "B": 9, "C": 9, "D": 14, "E": 46, "F": 12, "G": 12, "H": 26, "J": 8}.items():
            b.width(c, w)
        sheet.Columns.getByName("J").IsVisible = False

    def adjustments_logic(self):
        b = Builder(self.doc, self.adj)
        for k in range(BLOCKS):
            _, _, r1, r2, _ = self.block_rows(k)
            b.formulas(
                "F%d:H%d" % (r1, r2),
                [self.adjustment_formulas(r) for r in range(r1, r2 + 1)],
            )
            self.validate(self.adj, r1, r2)

    def adjustment_formulas(self, r):
        """F (from-year amount), G (to-year amount) and H (check) of an Adjustments row."""
        top = self.adj_top
        p = r - 1
        from_amount = '=IFS(OR($B{r}=""; $B{r}="none"; $D{r}=""); ""; TRUE(); $D{r})'.format(r=r)
        to_amount = (
            "=IFERROR(LET(\n"
            "  from; $B{r}; to; $C{r}; amount; $D{r};\n"
            "  IFS(\n"
            '    OR(to=""; to="none"; amount=""; from=""); "";\n'
            '    from="none"; amount;\n'
            "    TRUE(); amount*INDEX(PB_Inflator; MATCH(to; PB_Year; 0))/INDEX(PB_Inflator; MATCH(from; PB_Year; 0))));\n"
            '  "")'
        ).format(r=r)
        check = (
            "=LET(\n"
            "  item; $A{r}; from; $B{r}; to; $C{r}; amount; $D{r}; reason; $E{r};\n"
            "  row; IFERROR(MATCH(item; Exp_AdjLabel; 0); 0);\n"
            "  rfs; IFERROR(INDEX(RFS_Fv_Values; MATCH(INDEX(Exp_Key; row); RFS_Fv_Key; 0); MATCH(from; RFS_Fv_Years; 0)); 0);\n"
            "  moved_in; SUMIFS($G${t}:$G{p}; $A${t}:$A{p}; item; $C${t}:$C{p}; from);\n"
            "  moved_out; SUMIFS($F${t}:$F{p}; $A${t}:$A{p}; item; $B${t}:$B{p}; from);\n"
            "  IFS(\n"
            '    AND(item=""; from=""; to=""; amount=""; reason=""); "";\n'
            '    OR(item=""; from=""; to=""; amount=""; reason=""); "Missing entry";\n'
            '    row<=1; "Item not in Expenditures";\n'
            '    from=to; "Year from = Year to";\n'
            '    from="none"; "OK";\n'
            '    $F{r}>rfs+moved_in-moved_out+0.5; "More than the item holds in "&from;\n'
            '    TRUE(); "OK"))'
        ).format(r=r, p=p, t=top)
        return [from_amount, to_amount, check]

    def validate(self, sheet, r1, r2):
        def apply(rng, kind, formula, message):
            cells = sheet.getCellRangeByName(rng)
            v = cells.Validation
            v.Type = uno.Enum("com.sun.star.sheet.ValidationType", kind)
            # A relative formula is read from the range's first cell.
            v.setSourcePosition(cells.getCellByPosition(0, 0).CellAddress)
            v.Formula1 = formula
            v.IgnoreBlankCells = True
            v.ShowList = 1  # unsorted
            v.ShowErrorMessage = True
            v.ErrorAlertStyle = uno.Enum("com.sun.star.sheet.ValidationAlertStyle", "STOP")
            v.ErrorMessage = message
            cells.Validation = v

        apply("A%d:A%d" % (r1, r2), "LIST", "Exp_AdjLabel", "Pick an item from the list.")
        for c in "BC":
            apply("%s%d:%s%d" % (c, r1, c, r2), "LIST", "Adj_YearChoices",
                  'Pick "none" or a year of the RFS.')
        apply("D%d:D%d" % (r1, r2), "CUSTOM", "AND(ISNUMBER(D%d);D%d>0)" % (r1, r1),
              "The amount must be a number greater than 0, typed or given by a formula.")

    def user_adjustments(self):
        """The user's six rows in the block for the RFS's year 2 (2026)."""
        b = Builder(self.doc, self.adj)
        labels = {" ".join(l.split()): l for l in self.labels if l}
        r1 = self.block_rows(2)[2]
        for i, (row, amount) in enumerate(zip(self.user_rows, USER_AMOUNTS)):
            r = r1 + i
            item = " ".join(str(row[0]).split())
            if item not in labels:
                raise RuntimeError("Adjustments: %r is not an Expenditures label" % item)
            b.value("A%d" % r, labels[item])
            for c, v in (("B", row[1]), ("C", row[2])):
                b.value("%s%d" % (c, r), v if isinstance(v, float) else str(v).strip())
            if isinstance(amount, str):
                b.formula("D%d" % r, amount.format(prev=r - 1))
            else:
                b.value("D%d" % r, amount)
            b.value("E%d" % r, str(row[4]).strip())
            b.fix_heights(r, r)

    # -- Expenditures -------------------------------------------------------------------------

    def expenditures(self):
        doc, sheet = self.doc, self.exp
        b = Builder(doc, sheet)
        hdr, first, last, end = self.hdr, self.first, self.last, self.end

        self.fix_data(b)
        self.translate_notes(b)
        b.value("A1", "ADJUSTED EXPENDITURES")  # the mock-up's title

        cols = sheet.Columns
        # 1. The label column, J: the current-year block and the rest move one right.
        cols.insertByIndex(colnum("J") - 1, 1)
        # 2. Next year, year after and Reason for change (old O-U, now P-V) go; Working notes
        #    (old V) becomes P.
        cols.removeByIndex(colnum("P") - 1, 7)
        # 3. The RFS and Revised blocks (old W-BZ, now Q-BT) go.
        cols.removeByIndex(colnum("Q") - 1, 56)
        # 4. Q blank, R next year, S blank, T.. the adjusted RFS.
        adj_last = ADJ_FIRST + RFS_SPAN - 1
        cols.insertByIndex(colnum("Q") - 1, adj_last - colnum("Q") + 1)
        a0, a1 = col(ADJ_FIRST), col(adj_last)
        self.adj_cols = [col(ADJ_FIRST + k) for k in range(RFS_SPAN)]

        names = doc.NamedRanges
        for n in ("Exp_MidEst1", "Exp_LateEst1", "Exp_Current1", "Exp_MidEst2", "Exp_LateEst2",
                  "Exp_Current2", "Exp_Reason", "Exp_RFS", "Exp_RFSYears", "Exp_Revised",
                  "Exp_RevisedYears"):
            if names.hasByName(n):
                names.removeByName(n)
        define(doc, "Exp_AdjLabel", ref("Expenditures", "J%d:J%d" % (hdr, end)))
        define(doc, "Exp_NextYear", ref("Expenditures", "R%d:R%d" % (hdr, end)))
        define(doc, "Exp_Adjusted", ref("Expenditures", "%s%d:%s%d" % (a0, hdr, a1, end)))
        for n, c in (("Exp_Approved", "K"), ("Exp_Current", "O"), ("Exp_Notes", "P")):
            if ref("Expenditures", "%s%d:%s%d" % (c, hdr, c, end)) != names.getByName(n).Content:
                raise RuntimeError("%s is not where expected: %s" % (n, names.getByName(n).Content))

        # Clear whatever the inserted columns inherited.
        new = sheet.getCellRangeByName("Q1:%s%d" % (a1, end))
        new.clearContents(1 | 2 | 4 | 16 | 32 | 512)
        new.CellStyle = "Default"

        # J: the adjustment label.
        b.value("J%d" % hdr, "Adjustment label")
        b.formulas(
            "J%d:J%d" % (first, last),
            [[
                "=LET(\n"
                '  gl; TRIM($D{r}&""); element; TRIM($C{r}&""); name; TRIM($E{r}&"");\n'
                '  core; gl&" - "&IF(element<>""; element; name);\n'
                '  IFS(gl=""; ""; TRIM($F{r}&"")<>""; core&" (Spatac)"; TRUE(); core))'.format(r=r)
            ] for r in range(first, last + 1)],
        )
        b.cell("J%d:J%d" % (hdr, end)).CellStyle = "Default"
        b.style("J%d" % hdr, "Header")
        b.style("J%d:J%d" % (first, last), "Calculated")
        b.wrap("J%d:J%d" % (first, last))

        # The adjusted block's years are numbers in row 6, which formulas read; they are green
        # (calculated) and invisible. The headings in row 7 show them as text: row 7 is inside the
        # block's names, so a number there would be added in by every column sum (design 14.12).
        right = uno.Enum("com.sun.star.table.CellHoriJustify", "RIGHT")
        b.formula("%s6" % a0, "=RFS_StudyYear")
        for k in range(1, RFS_SPAN):
            b.formula("%s6" % self.adj_cols[k], "=%s6+1" % self.adj_cols[k - 1])
        for c in self.adj_cols:
            b.formula("%s%d" % (c, hdr), '=""&%s$6' % c)
        b.style("%s6:%s6" % (a0, a1), "Calculated")
        b.fmt("%s6:%s6" % (a0, a1), ";;;")
        # The comment and the block's description go in the spacer before the block (S here, T
        # once the label's spacer is inserted): it is never folded, unlike the block's past years.
        sheet.Annotations.insertNew(
            b.cell("S6").CellAddress,
            "Do not modify the cells to the right of this one in this row: they hold the years of "
            "the adjusted block, which the formulas read.",
        )
        define(doc, "Exp_AdjustedYears", ref("Expenditures", "%s6:%s6" % (a0, a1)))
        year_of = lambda c: "%s$6" % c

        # R: next year, from the adjusted block. Its heading row 2 is just the year.
        next_col = "Cfg_BaseYear+1"
        b.formula("R2", "=Cfg_BaseYear+1")
        b.fmt("R2", YEAR)
        b.bold("R2")
        b.formula("R%d" % hdr, '=(Cfg_BaseYear+1)&" Adjusted"')
        b.formula("R3", "=INDEX($%s$3:$%s$3;MATCH(%s;Exp_AdjustedYears;0))" % (a0, a1, next_col))
        b.formula("R4", "=SUM(R%d:R%d)" % (hdr, end))
        b.formula("R5", "=R3-R4")
        b.formulas(
            "R%d:R%d" % (first, last),
            [["=INDEX($%s%d:$%s%d;1;MATCH(%s;Exp_AdjustedYears;0))" % (a0, r, a1, r, next_col)]
             for r in range(first, last + 1)],
        )
        b.style("R%d:R%d" % (first, last), "Calculated")
        box(b.cell("R2:R5"))

        # The adjusted RFS block.
        # Two leading spaces keep the text clear of the next-year box to its left.
        b.value("S2", "  RFS amounts by year, with the Adjustments tab applied (amounts that "
                      "differ from the RFS are shaded orange)")
        b.bold("S2")
        b.cell("S2").IsTextWrapped = False
        # Legends for rows 3-5, right-justified next to the figures (in the notes column).
        for r, text in ((3, "RFS total"), (4, "Adjusted total"), (5, "Variance vs RFS")):
            b.value("P%d" % r, text)
            b.cell("P%d" % r).HoriJustify = right
        # Variances: negative is unfavourable (more spending than the RFS), in red, as on
        # Medium_Term_Cashflow.
        b.value("C5", "Variance vs RFS (more spending than the RFS is negative, in red)")
        for c in self.adj_cols:
            b.formula("%s3" % c, "=IFERROR(-INDEX(RFS_Fv_Totals;1;MATCH(%s;RFS_Fv_Years;0));0)" % year_of(c))
            b.formula("%s4" % c, "=SUM(%s%d:%s%d)" % (c, hdr, c, end))
            b.formula("%s5" % c, "=%s3-%s4" % (c, c))
        # An amount is blank when the RFS has none and no adjustment touches it; otherwise it
        # shows, a deliberate zero as 0 (design 13.4).
        b.formulas(
            "%s%d:%s%d" % (a0, first, a1, last),
            [[
                "=LET(\n"
                "  year; {y};\n"
                "  rfs; IFERROR(INDEX(RFS_Fv_Values; MATCH($G{r}; RFS_Fv_Key; 0); MATCH(year; RFS_Fv_Years; 0)); 0);\n"
                "  touched; COUNTIFS(Adj_Item; $J{r}; Adj_To; year)+COUNTIFS(Adj_Item; $J{r}; Adj_From; year);\n"
                "  added; SUMIFS(Adj_ToAmount; Adj_Item; $J{r}; Adj_To; year);\n"
                "  removed; SUMIFS(Adj_FromAmount; Adj_Item; $J{r}; Adj_From; year);\n"
                "  IFS(\n"
                '    AND(rfs=0; OR($J{r}=""; touched=0)); "";\n'
                '    $J{r}=""; rfs;\n'
                "    TRUE(); rfs+added-removed))".format(r=r, y=year_of(c))
                for c in self.adj_cols
            ] for r in range(first, last + 1)],
        )
        b.formula(
            "O5",
            "=INDEX($%s$3:$%s$3;MATCH(Cfg_BaseYear;Exp_AdjustedYears;0))-O4" % (a0, a1),
        )
        self.highlight_changes(sheet, a0, a1)

        # Formats.
        b.style("Q%d:%s%d" % (hdr, a1, hdr), "Header")
        b.style("Q%d:Q%d" % (hdr, hdr), "Default")
        b.style("S%d:S%d" % (hdr, hdr), "Default")
        b.wrap("J%d:%s%d" % (hdr, a1, hdr))
        b.cell("%s%d:%s%d" % (a0, hdr, a1, hdr)).HoriJustify = right
        b.fmt("R3:%s4" % a1, MONEY_ZERO)  # totals: a zero shows 0 (design 15.6)
        for rng in ("O5", "R5", "%s5:%s5" % (a0, a1)):
            b.fmt(rng, MONEY_RED)
        b.bold("R4")
        b.bold("%s4:%s4" % (a0, a1))
        b.fmt("R%d:R%d" % (first, last), MONEY_ZERO)
        b.fmt("%s%d:%s%d" % (a0, first, a1, last), MONEY_ZERO)
        b.fill("A%d:%s%d" % (end, a1, end), SEPARATOR_COLOR)
        for c, w in {"J": 30, "Q": 1.5, "R": 13, "S": 1.5}.items():
            b.width(c, w)
        for c in self.adj_cols:
            b.width(c, 10)

        base = int(self.doc.NamedRanges.getByName("Cfg_BaseYear").ReferredCells.getValue())
        study = int(self.doc.NamedRanges.getByName("RFS_StudyYear").ReferredCells.getValue())

        # The labels as the formulas give them, for the user's rows.
        doc.calculateAll()
        self.labels = [r[0] for r in b.cell("J%d:J%d" % (first, last)).getDataArray()]
        dupes = {l for l in self.labels if l and self.labels.count(l) > 1}
        if dupes:
            raise RuntimeError("Expenditures: duplicate adjustment labels %s" % sorted(dupes))

        # 3.9's 2028 amount differs from the golden source (design 14.4, G32).
        row39 = self.row_of(b, "3.9")
        y2028 = self.adj_cols[2028 - study]
        sheet.Annotations.insertNew(b.cell("%s%d" % (y2028, row39)).CellAddress, COMMENT_39)

        # A narrow spacer after the label (K), so that the label's folded group stays a group of
        # its own: in .xlsx, groups that touch merge into one. Everything from the current year on
        # moves one column right, with its formulas, names and formats (design 14.12).
        sheet.Columns.insertByIndex(colnum("K") - 1, 1)
        spacer = b.cell("K1:K%d" % end)
        spacer.clearContents(1 | 2 | 4 | 16 | 32 | 512)
        spacer.CellStyle = "Default"
        b.fill("K%d" % end, SEPARATOR_COLOR)
        b.width("K", 1.5)
        a0, a1 = col(ADJ_FIRST + 1), col(ADJ_FIRST + RFS_SPAN)
        self.adj_cols = [col(ADJ_FIRST + 1 + k) for k in range(RFS_SPAN)]
        if b.cell("%s6" % a0).getValue() != study or b.cell("%s7" % LABEL).getString() != "Adjustment label":
            raise RuntimeError("Expenditures: the spacer did not move the columns as expected")
        # Approved and Actuals to date show a zero as 0, not blank (design 15.6).
        b.fmt("L%d:M%d" % (first, last), MONEY_ZERO)

        # Groups, folded: the identity helpers (E-H), the label (J) and the past years of the
        # adjusted block (one more is folded at each rollover, design 14.6). The current year's
        # inputs are not grouped.
        sheet.clearOutline()
        b.group_cols("E", "H", collapse=True)
        b.group_cols("J", "J", collapse=True)
        if base > study:
            b.group_cols(a0, self.adj_cols[base - study - 1], collapse=True)

        # The AutoFilter covers every column.
        udb = doc.getPropertyValue("UnnamedDatabaseRanges")
        udb.setByTable(b.cell("A%d:%s%d" % (hdr, a1, end)).RangeAddress)
        udb.getByTable(sheet.RangeAddress.Sheet).AutoFilter = True

    def row_of(self, b, item, element=None):
        data = b.cell("B%d:C%d" % (self.first, self.last)).getDataArray()
        rows = [
            self.first + i for i, (it, el) in enumerate(data)
            if str(it).strip() == item and (element is None or " ".join(str(el).split()) == element)
        ]
        if len(rows) != 1:
            raise RuntimeError("Expenditures: item %s %s found %d times" % (item, element, len(rows)))
        return rows[0]

    def fix_data(self, b):
        """The label fixes of design 14.6 (G23)."""
        r = self.row_of(b, "4.1a")
        if b.cell("D%d" % r).getString() != "9250.101":
            raise RuntimeError("Expenditures D%d is not 9250.101" % r)
        b.cell("D%d" % r).setString("9250.1010")
        self.report.append("Expenditures D%d (4.1a): GL account 9250.101 -> 9250.1010" % r)
        r = self.row_of(b, "18.24a")
        name = b.cell("E%d" % r).getString()
        if "\n" not in name:
            raise RuntimeError("Expenditures E%d has no line break" % r)
        b.cell("E%d" % r).setString(" ".join(name.split()))
        self.report.append("Expenditures E%d (18.24a): line break removed from the GL account name" % r)

    def translate_notes(self, b):
        """Old V, still in place: the column letters the notes name (design 14.8, G29)."""
        self.note_changes = []
        rng = b.cell("V%d:V%d" % (self.first, self.last))
        data = [list(r) for r in rng.getDataArray()]
        items = [str(r[0]).strip() for r in b.cell("B%d:B%d" % (self.first, self.last)).getDataArray()]
        fixed = set()
        for i, (text,) in enumerate(data):
            if not isinstance(text, str) or not text:
                continue
            new = translate_note(text)
            if items[i] in NOTE_FIXES:
                old_part, new_part = NOTE_FIXES[items[i]]
                if new.count(old_part) != 1:
                    raise RuntimeError("note of %s: %r not found once" % (items[i], old_part))
                new = new.replace(old_part, new_part)
                fixed.add(items[i])
            if new != text:
                self.note_changes.append((self.first + i, text, new))
                b.cell("V%d" % (self.first + i)).setString(new)
        if fixed != set(NOTE_FIXES):
            raise RuntimeError("notes not corrected: %s" % sorted(set(NOTE_FIXES) - fixed))

    def highlight_changes(self, sheet, a0, a1):
        rng = sheet.getCellRangeByName("%s%d:%s%d" % (a0, self.first, a1, self.last))
        cf = rng.ConditionalFormat
        cf.clear()

        def pv(name, value):
            p = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            p.Name, p.Value = name, value
            return p

        cf.addNew((
            pv("Operator", uno.Enum("com.sun.star.sheet.ConditionOperator", "FORMULA")),
            pv("Formula1",
               "ABS(N({c}{r})-IFERROR(INDEX(RFS_Fv_Values;MATCH($G{r};RFS_Fv_Key;0);"
               "MATCH({c}$6;RFS_Fv_Years;0));0))>0.5".format(c=a0, r=self.first)),
            pv("StyleName", "Changed"),
            pv("SourcePosition", rng.getCellByPosition(0, 0).CellAddress),
        ))
        rng.ConditionalFormat = cf

    # -- Medium_Term_Cashflow -----------------------------------------------------------------

    def mtc(self):
        sheet = self.doc.Sheets.getByName("Medium_Term_Cashflow")
        b = Builder(self.doc, sheet)
        if b.cell("A24").getString() != "Minimum closing cash balance (Settings, inflated)":
            raise RuntimeError("Medium_Term_Cashflow A24 is not the minimum row")
        # A blank row after the adjustment row; the minimum row is relabelled (design 14.3).
        sheet.Rows.insertByIndex(23, 1)
        blank = b.cell("A24:G24")
        blank.clearContents(1 | 2 | 4 | 16 | 32 | 512)
        blank.CellStyle = "Default"
        b.cell("H24").CellBackColor = b.cell("H23").CellBackColor
        b.value("A25", "Minimum closing cash balance")
        b.cell("A25:G25").CharWeight = 100.0
        years = ["B", "C", "D", "E", "F", "G"]
        for c in years:
            b.formula(
                "%s9" % c,
                "=LET(\n"
                "  year; {c}$3;\n"
                "  IF(year=Cfg_BaseYear; -SUM(Exp_Current); -SUM(INDEX(Exp_Adjusted; 0; MATCH(year; Exp_AdjustedYears; 0)))))".format(c=c),
            )
            b.formula("%s29" % c, "=INDEX(RFS_Fv_Totals;1;MATCH(%s$3;RFS_Fv_Years;0))" % c)
        if b.cell("A29").getString() != "RFS expenditures":
            raise RuntimeError("Medium_Term_Cashflow A29 is not the RFS row")

    # -- Projected_Balances -------------------------------------------------------------------

    def projected(self):
        """Rebuilt (design 15.2-15.4): the RFS block recomputed from the RFS's inputs and checked
        against RFS_cashflow, beside a what-if that always starts in the current year. The tab
        is cleared and rebuilt in place, so that the names other tabs read keep working."""
        doc = self.doc
        sheet = doc.Sheets.getByName("Projected_Balances")
        b = Builder(doc, sheet)
        # Header row 13, years in rows 14-63: the blank row 11 is gone (design 15.7).
        hdr, first, last = 13, 14, 63
        last30 = first + RFS_SPAN - 1

        # The user's inputs, kept: opening balance, the "through" year, the catch-up increases and
        # the subsequent increase (a formula by default).
        expected = {
            "A5": "Opening balance of the current year (Current year start only)",
            "A8": "Contributions follow the RFS schedule through",
            "A12": "Subsequent annual increase",
        }
        for a, text in expected.items():
            if b.cell(a).getString() != text:
                raise RuntimeError("Projected_Balances %s is not %r" % (a, text))
        opening = b.cell("B5").getValue()
        through = b.cell("B8").getValue()
        catch_up = [b.cell("B%d" % r).getValue() for r in (9, 10, 11)]
        subsequent = b.cell("B12").getFormula()
        if not remove_comment(sheet, "B5"):  # its instructions are in the docs (design 14.9)
            raise RuntimeError("no comment on Projected_Balances!B5")

        everything = b.cell("A1:AF200")
        everything.Validation = sheet.getCellRangeByName("AF200").Validation  # none
        everything.clearContents(1 | 2 | 4 | 8 | 16 | 32 | 64 | 256 | 512)
        everything.CellStyle = "Default"
        cf = everything.ConditionalFormat
        cf.clear()
        everything.ConditionalFormat = cf
        sheet.clearOutline()
        names = doc.NamedRanges
        for n in ("PB_Start", "PB_ExpBasis", "PB_ContribBasis", "PB_StartBalance", "PB_RFSCheck"):
            if names.hasByName(n):
                names.removeByName(n)

        def name(n, a):
            # Names other tabs read (PB_Year, PB_Inflator, PB_Balance, PB_RFSBalance) are moved,
            # not re-created, so the formulas that use them keep them.
            content = ref("Projected_Balances", a)
            if names.hasByName(n):
                names.getByName(n).setContent(content)
            else:
                define(doc, n, content)

        b.value("A1", "PROJECTED RESERVE FUND BALANCES").CellStyle = "Title"
        b.value(
            "A2",
            "The RFS's projection recomputed from its inputs (columns B-M), checked against "
            "RFS_cashflow (column I), and a what-if projection (columns O-Z, inputs in P4:P10).",
        )

        # The what-if's inputs, above the what-if (design 15.5). The labels are right-justified
        # in O, so that they run left over the empty cells before it.
        right = uno.Enum("com.sun.star.table.CellHoriJustify", "RIGHT")
        b.value("O4", "What-if starts in")
        b.formula("P4", "=Cfg_BaseYear")
        b.style("P4", "Calculated")
        b.fmt("P4", YEAR)
        name("PB_StartYear", "P4")
        inputs = [
            (5, "Opening balance of the current year", opening, MONEY_RED, "PB_OpeningBalance"),
            (6, "Contribution increases follow the RFS schedule through", through, YEAR,
             "PB_ContribFixedThrough"),
            (7, "Contribution increase, 1st year after", catch_up[0], PERCENT, None),
            (8, "Contribution increase, 2nd year after", catch_up[1], PERCENT, None),
            (9, "Contribution increase, 3rd year after", catch_up[2], PERCENT, None),
            (10, "Subsequent annual increase", subsequent, PERCENT, "PB_SubsequentIncrease"),
        ]
        for r, label, v, fmt, n in inputs:
            b.value("O%d" % r, label)
            if isinstance(v, str):
                b.formula("P%d" % r, v)
            else:
                b.value("P%d" % r, v)
            b.style("P%d" % r, "Input")
            b.fmt("P%d" % r, fmt)
            if n:
                name(n, "P%d" % r)
        b.cell("O4:O10").HoriJustify = right
        name("PB_CatchUp", "P7:P9")

        # A summary above each block: over the first 30 years of the RFS, each margin's lowest
        # value and its first negative one, with their years (design 15.5).
        self.margin_summary(b, "B", "RFS, first 30 years", ("J", "K", "L"), first, last30)
        self.margin_summary(b, "R", "What-if, first 30 years", ("W", "X", "Y"), first, last30)
        name("PB_RFSLowest", "B5:G7")
        name("PB_Lowest", "R5:W7")

        # The year table.
        b.value("B%d" % (hdr - 1), "RFS (recomputed from RFS_cashflow inputs)")
        b.value("O%d" % (hdr - 1), "What-if")
        b.bold("A%d:Z%d" % (hdr - 1, hdr - 1))
        b.value("H%d" % (hdr - 1), "Years that differ")
        b.cell("H%d" % (hdr - 1)).HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", "RIGHT")
        b.formula("I%d" % (hdr - 1), '=COUNTIF(I%d:I%d;">0.5")+COUNTIF(I%d:I%d;"<-0.5")'
                  % (first, last, first, last))
        b.fmt("I%d" % (hdr - 1), "[RED]0;-0;0")  # a count: red when not zero
        heads = {
            "A": "Year",
            "B": "RFS opening balance", "C": "RFS expected costs", "D": "RFS contribution increase",
            "E": "RFS contributions", "F": "RFS interest rate", "G": "RFS interest",
            "H": "RFS closing balance", "I": "Difference from RFS_cashflow balance",
            "J": "RFS margin over minimum", "K": "RFS margin over inflated minimum",
            "L": "RFS PV margin over minimum", "M": "RFS contribution per unit per month",
            "O": "Opening balance", "P": "Expenditures", "Q": "Contribution increase",
            "R": "Contributions", "S": "Interest rate", "T": "Interest", "U": "Closing balance",
            "V": "Difference from RFS balance", "W": "Margin over minimum",
            "X": "Margin over inflated minimum", "Y": "PV margin over minimum",
            "Z": "Contribution per unit per month", "AB": "Inflator",
        }
        for c, h in heads.items():
            b.value("%s%d" % (c, hdr), h)
        b.style("A%d:AB%d" % (hdr, hdr), "Header")
        b.style("N%d" % hdr, "Default")
        b.style("AA%d" % hdr, "Default")
        b.wrap("A%d:AB%d" % (hdr, hdr))
        b.height(hdr, 42)

        rows = []
        for r in range(first, last + 1):
            p = r - 1
            top = r == first
            row = [
                # A: Year
                "=RFS_StudyYear" if top else "=$A%d+1" % p,
                # B-M: the RFS, recomputed. Read: costs, rates, opening balance, present
                # contribution. Opening + costs (negative) + contributions + interest = closing.
                "=RFS_OpeningBalance" if top else "=$H%d" % p,
                "=IFERROR(INDEX(RFS_Cf_Cost;MATCH($A{r};RFS_Cf_Year;0));0)",
                "" if top else (
                    "=LET(\n"
                    "  n; $A{r}-RFS_StudyYear;\n"
                    "  IF(n<=3; INDEX(RFS_CatchUpIncreases; n); RFS_SubsequentIncrease))"
                ),
                "=RFS_PresentContribution" if top else "=$E%d*(1+$D{r})" % p,
                "=IF($A{r}<RFS_StudyYear+Cfg_ShortInterestYears;RFS_ShortInterest;RFS_LongInterest)",
                "=$B{r}*$F{r}",
                "=$B{r}+$C{r}+$E{r}+$G{r}",
                '=IFERROR($H{r}-INDEX(RFS_Cf_Balance;MATCH($A{r};RFS_Cf_Year;0));"")',
                "=$H{r}-RFS_MinBalance",
                "=$H{r}-RFS_MinBalance*$AB{r}",
                "=$H{r}/$AB{r}-RFS_MinBalance",
                "=$E{r}/12/RFS_Units",
                "",  # N: spacer
                # O-Z: the what-if, from the start year on; before it every what-if cell is blank.
                '=IF($A{r}<PB_StartYear;"";PB_OpeningBalance)' if top else
                '=IFS($A{r}<PB_StartYear;"";$A{r}=PB_StartYear;PB_OpeningBalance;TRUE();$U%d)' % p,
                # P: Expenditures: the current year's estimates, then the adjusted RFS, then the
                # RFS beyond the adjusted block.
                "=LET(\n"
                "  year; $A{r};\n"
                "  adjusted_column; MATCH(year; Exp_AdjustedYears; 0);\n"
                "  IFS(\n"
                '    $O{r}=""; "";\n'
                "    year=Cfg_BaseYear; -SUM(Exp_Current);\n"
                "    ISNA(adjusted_column); $C{r};\n"
                "    TRUE(); -SUM(INDEX(Exp_Adjusted; 0; adjusted_column))))",
                # Q: the RFS's increase through the chosen year, then the catch-up increases,
                # then the subsequent one.
                "=LET(\n"
                "  years_after; $A{r}-PB_ContribFixedThrough;\n"
                "  IFS(\n"
                '    OR($O{r}=""; $A{r}=RFS_StudyYear); "";\n'
                "    years_after<=0; $D{r};\n"
                "    years_after<=3; INDEX(PB_CatchUp; years_after);\n"
                "    TRUE(); PB_SubsequentIncrease))",
                # R: last year's contributions (the RFS block's before the what-if starts), grown.
                '=IF($O{r}="";"";RFS_PresentContribution)' if top else
                '=IFS($O{r}="";"";$A{r}=RFS_StudyYear;RFS_PresentContribution;'
                'TRUE();IF($R%d="";$E%d;$R%d)*(1+$Q{r}))' % (p, p, p),
                '=IF($O{r}="";"";$F{r})',
                '=IF($O{r}="";"";$O{r}*$S{r})',
                '=IF($O{r}="";"";$O{r}+$P{r}+$R{r}+$T{r})',
                '=IF($U{r}="";"";$U{r}-$H{r})',
                '=IF($U{r}="";"";$U{r}-RFS_MinBalance)',
                '=IF($U{r}="";"";$U{r}-RFS_MinBalance*$AB{r})',
                '=IF($U{r}="";"";$U{r}/$AB{r}-RFS_MinBalance)',
                '=IF($R{r}="";"";$R{r}/12/RFS_Units)',
                "",  # AA: spacer
                # AB: the inflation factor (design 14.5).
                "=LET(\n"
                "  n; $A{r}-RFS_StudyYear;\n"
                "  (1+RFS_ShortInflation)^MIN(n; Cfg_ShortInflationYears)*(1+RFS_LongInflation)^MAX(n-Cfg_ShortInflationYears; 0))",
            ]
            rows.append([f.format(r=r) if f else "" for f in row])
        b.formulas("A%d:AB%d" % (first, last), rows)

        b.style("B%d:M%d" % (first, last), "Calculated")
        b.style("O%d:Z%d" % (first, last), "Calculated")
        b.fmt("A%d:A%d" % (first, last), YEAR)
        for c in ("B", "C", "E", "G", "H", "I", "J", "K", "L",
                  "O", "P", "R", "T", "U", "V", "W", "X", "Y"):
            b.fmt("%s%d:%s%d" % (c, first, c, last), MONEY_RED)
        for c in "DFQS":
            b.fmt("%s%d:%s%d" % (c, first, c, last), PERCENT)
        for c in "MZ":
            b.fmt("%s%d:%s%d" % (c, first, c, last), MONEY_CENTS_RED)
        b.fmt("AB%d:AB%d" % (first, last), "0.0000")
        # Years past the 30-year horizon (design 15.4).
        b.fill("A%d:Z%d" % (last30 + 1, last), BEYOND_COLOR)
        self.flag_check(sheet, "I%d:I%d" % (first, last))

        b.width("A", 8)  # years only: the inputs' labels are in O
        for k in range(colnum("B"), colnum("AB") + 1):
            b.width(col(k), 13)
        b.width("N", 2)
        b.width("AA", 2)
        b.group_cols("AB", "AB", collapse=True)
        ctrl = doc.CurrentController
        ctrl.setActiveSheet(sheet)
        ctrl.freezeAtPosition(1, hdr)

        for n, c in (
            ("PB_Year", "A"), ("PB_RFSOpening", "B"), ("PB_RFSCost", "C"),
            ("PB_RFSIncrease", "D"), ("PB_RFSContrib", "E"), ("PB_RFSRate", "F"),
            ("PB_RFSInterest", "G"), ("PB_RFSBalance", "H"), ("PB_RFSDifference", "I"),
            ("PB_RFSMargin", "J"), ("PB_RFSMarginInflated", "K"), ("PB_RFSMarginPV", "L"),
            ("PB_RFSPerUnitMonth", "M"),
            ("PB_Opening", "O"), ("PB_Cost", "P"), ("PB_Increase", "Q"),
            ("PB_Contributions", "R"), ("PB_Rate", "S"), ("PB_Interest", "T"),
            ("PB_Balance", "U"), ("PB_Variance", "V"), ("PB_Margin", "W"),
            ("PB_MarginInflated", "X"), ("PB_MarginPV", "Y"), ("PB_PerUnitMonth", "Z"),
            ("PB_Inflator", "AB"),
        ):
            name(n, "%s%d:%s%d" % (c, first, c, last))

    def margin_summary(self, b, at, title, margins, first, last30):
        """Label at `at` (running over the next column); then Lowest, Year, First negative, Year."""
        c0 = colnum(at)
        low, low_year, neg, neg_year = (col(c0 + k) for k in (2, 3, 4, 5))
        years = "$A$%d:$A$%d" % (first, last30)
        b.values("%s4:%s4" % (at, neg_year), [[title, "", "Lowest", "Year", "First negative", "Year"]])
        b.style("%s4:%s4" % (at, neg_year), "Header")
        labels = ("Margin over minimum", "Margin over inflated minimum", "PV margin over minimum")
        for i, (label, m) in enumerate(zip(labels, margins)):
            r = 5 + i
            span = "%s%d:%s%d" % (m, first, m, last30)
            b.value("%s%d" % (at, r), label)
            b.formula("%s%d" % (low, r), "=MIN(%s)" % span)
            b.formula("%s%d" % (low_year, r), "=INDEX(%s;MATCH(%s%d;%s;0))" % (years, low, r, span))
            b.formula(
                "%s%d" % (neg_year, r),
                '=IF(COUNTIF({s};"<0")=0;"none";MINIFS({y};{s};"<0"))'.format(s=span, y=years),
            )
            b.formula(
                "%s%d" % (neg, r),
                '=IF({ny}{r}="none";"";INDEX({s};MATCH({ny}{r};{y};0)))'.format(ny=neg_year, r=r, s=span, y=years),
            )
        b.fmt("%s5:%s7" % (low, low), MONEY_RED)
        b.fmt("%s5:%s7" % (neg, neg), MONEY_RED)
        for c in (low_year, neg_year):
            b.fmt("%s5:%s7" % (c, c), YEAR)
            b.cell("%s5:%s7" % (c, c)).HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", "RIGHT")

    def flag_check(self, sheet, address):
        """Shades a year whose recomputed RFS balance differs from RFS_cashflow's by more than
        0.5 (design 15.2)."""
        rng = sheet.getCellRangeByName(address)
        cf = rng.ConditionalFormat
        cf.clear()

        def pv(n, value):
            p = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            p.Name, p.Value = n, value
            return p

        first = rng.getCellByPosition(0, 0)
        cf.addNew((
            pv("Operator", uno.Enum("com.sun.star.sheet.ConditionOperator", "FORMULA")),
            pv("Formula1", "ABS(N(%s))>0.5" % address.split(":")[0]),
            pv("StyleName", "Changed"),
            pv("SourcePosition", first.CellAddress),
        ))
        rng.ConditionalFormat = cf

    # -- README -------------------------------------------------------------------------------

    def readme(self):
        sheet = self.doc.Sheets.getByName("README")
        b = Builder(self.doc, sheet)
        if not b.cell("A6").getString().startswith("Expenditures"):
            raise RuntimeError("README A6 is not the Expenditures line")
        b.value(
            "A6",
            "Expenditures — every RFS element, with the current year's approved amounts, actuals, "
            "estimates and working notes, and the RFS amounts by year with the Adjustments applied "
            "(amounts that differ from the RFS are highlighted).",
        )
        sheet.Rows.insertByIndex(6, 1)
        b.value(
            "A7",
            "Adjustments — the Board's changes to the RFS, each with its reason: amounts moved "
            "between years, added or removed.",
        )
        b.cell("A7").CellStyle = b.cell("A6").CellStyle
        b.cell("A7").IsTextWrapped = True
        legend = [i for i in range(14, 20) if b.cell("A%d" % i).getString().startswith("Grey")]
        if len(legend) != 1:
            raise RuntimeError("README: the Grey legend line was not found")
        g = legend[0]
        sheet.Rows.insertByIndex(g, 1)
        b.value("A%d" % (g + 1), "Orange — on Expenditures, an amount that differs from the RFS "
                                 "because of the Adjustments tab.")
        b.cell("A%d" % (g + 1)).CellStyle = b.cell("A%d" % g).CellStyle
        b.cell("A%d" % (g + 1)).IsTextWrapped = True
        settings = [i for i in range(8, 16) if b.cell("A%d" % i).getString().startswith("Settings")]
        if len(settings) != 1:
            raise RuntimeError("README: the Settings line was not found")
        b.value("A%d" % settings[0], "Settings — the base year, the minimum target cash balance, the "
                "rate assumed for planned new investments, and the RFS's years of short-term "
                "inflation and of short-term interest.")
        projected = [i for i in range(8, 16) if b.cell("A%d" % i).getString().startswith("Projected_Balances")]
        if len(projected) != 1:
            raise RuntimeError("README: the Projected_Balances line was not found")
        b.value("A%d" % projected[0], "Projected_Balances — the reserve fund balance to the end of "
                "the RFS: the RFS's own projection, recomputed from its inputs and checked against "
                "RFS_cashflow, and a what-if with Spatac's adjustments and contribution increases.")
        # A legend line for the years past the RFS's horizon, after Orange (design 15.4).
        sheet.Rows.insertByIndex(g + 1, 1)
        b.value("A%d" % (g + 2), "Light grey — on Projected_Balances, a year past the RFS's 30-year "
                                 "horizon.")
        b.cell("A%d" % (g + 2)).CellStyle = b.cell("A%d" % g).CellStyle
        b.cell("A%d" % (g + 2)).IsTextWrapped = True
        b.fix_heights(6, 7)
        b.fix_heights(settings[0], settings[0])
        b.fix_heights(projected[0], projected[0])
        b.fix_heights(g + 1, g + 2)

    # -- comments and tab order ---------------------------------------------------------------

    def instruction_comments(self):
        """Claude's instruction comments move into the docs (design 14.9)."""
        sheets = self.doc.Sheets
        # Projected_Balances B5's goes when the tab is rebuilt (projected).
        for sheet, address in (("Medium_Term_Cashflow", "H3"), ("RFS_info", "K1"),
                               ("RFS_fvalues", "BC1")):
            if not remove_comment(sheets.getByName(sheet), address):
                raise RuntimeError("no comment on %s!%s" % (sheet, address))

    def no_hidden_zeros(self):
        """No cell on a Spatac tab hides a zero (design 15.6). The source still has the old
        hide-zero format (#,##0;-#,##0;""), on the Schedule_of_Investments cash-flow grid: it
        becomes #,##0."""
        formats, locale = self.doc.NumberFormats, self.doc.CharLocale
        plain = formats.queryKey(MONEY_ZERO, locale, False)
        if plain == -1:
            plain = formats.addNew(MONEY_ZERO, locale)
        hides = {}

        def hides_zero(key):
            if key not in hides:
                hides[key] = formats.getByKey(key).FormatString.endswith(';""')
            return hides[key]

        count = 0
        for sheet in self.doc.Sheets:
            if sheet.Name.startswith("RFS_"):
                continue
            cur = sheet.createCursor()
            cur.gotoEndOfUsedArea(False)
            a = cur.RangeAddress
            for r in range(a.EndRow + 1):
                for c in range(a.EndColumn + 1):
                    cell = sheet.getCellByPosition(c, r)
                    if hides_zero(cell.NumberFormat):
                        cell.NumberFormat = plain
                        count += 1
        self.report.append("%d cells no longer hide a zero" % count)

    def order_and_protect(self):
        sheets = self.doc.Sheets
        order = [
            "README", "Expenditures", "Adjustments", "Medium_Term_Cashflow",
            "Schedule_of_Investments", "Projected_Balances", "RFS_info", "RFS_fvalues",
            "RFS_cashflow", "Settings",
        ]
        for i, name in enumerate(order):
            sheets.moveByName(name, i)
        for name in ("README", "Settings", "Expenditures", "Adjustments", "Medium_Term_Cashflow",
                     "Schedule_of_Investments", "Projected_Balances"):
            sheet = sheets.getByName(name)
            if sheet.isProtected():
                sheet.unprotect("")
            sheet.protect("")
        self.doc.CurrentController.setActiveSheet(sheets.getByName("README"))


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("source")
    ap.add_argument("mockup")
    ap.add_argument("output")
    ap.add_argument("--notes", help="write the translated working notes to this file for review")
    args = ap.parse_args()
    with Session() as s:
        r = Restructure(s, args.source, args.mockup, args.output)
        report = r.run()
    print("Wrote %s" % args.output)
    for line in report:
        print("- " + line)
    print("- %d working notes now name their columns by heading" % len(r.note_changes))
    if args.notes:
        with open(args.notes, "x", encoding="utf-8") as f:
            for row, old, new in r.note_changes:
                f.write("Row %d\n  before: %s\n  after:  %s\n\n" % (row, old, new))


if __name__ == "__main__":
    sys.exit(main())
