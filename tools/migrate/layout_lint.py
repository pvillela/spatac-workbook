"""Layout checks that apply to every tab of any build (design-history.md section 17.2).

Usage:

    /opt/libreoffice26.2/program/python tools/migrate/layout_lint.py WORKBOOK.xlsx

The workbook is opened read-only, as a copy. The checks, on every tab:

- **Protection**: every tab is protected; every unlocked cell is a yellow input cell (it may
  hold the user's own formula).
- **Colours**: a conditional format that sets a font colour is not overridden by a colour code
  in a cell's number format (`[RED]`), which wins over the conditional format.
- **Orphans**: no conditional format covers only empty cells; every discrepancy count
  (`=SUMPRODUCT(--(ABS(`) has the Alarm format.
- **Display**: no number shows as "-0" (a rounding remainder).

and, on every tab but the RFS tabs (their layout is the vendor's, design 17.6):

- **Heights**: a row with an input cell has an automatic height, so that it grows as the user
  types (design 17.3); no wrapped text in a hidden column sets an automatic row's height.
- **Fit**: every cell's content fits its cell (or merged area). A fixed-height row must fit its
  wrapped text even when the text renders `WIDER` wider, as it does in Excel or at 110-120%
  zoom; an automatic row is fitted by the program itself, so it gets no margin. Text on one
  line must fit the width with the margin, or overflow into empty neighbours; numbers must fit
  the width. Text typed into an input cell or a number column is the user's and is not
  measured for width.
- **Filter buttons**: in an AutoFilter's header row, the drop-down button covers the right side
  of each cell; its width (`FILTER_BUTTON`) is not available to the text.
- **Breaks**: wrapped text never breaks inside a word; in a fixed-height row (a heading) not at
  a hyphen either ("Late-/year"). Lines are filled as the programs fill them.

`lint(doc)` returns the problems as strings; the verifiers call it on their own copy, since the
checks add a scratch tab and unprotect the tabs. The build script uses the same measurements
(`needed_width`, `needed_height`) to size columns and rows.
"""

import re
import sys

import uno  # noqa: F401  (registers the com.sun.star import hook)
from com.sun.star.table import CellAddress

from uno_session import Session

WIDER = 1.1  # how much wider text may render in Excel or at 110-120% zoom
SLACK = 10  # 1/100 mm: rounding in LibreOffice's fitted sizes
FILTER_BUTTON = 500  # 1/100 mm: an AutoFilter drop-down button (about 17 pixels at 100%)
INPUT_FILL = 0xFFF2CC
ALARM_FILL = 0xC00000
COLOUR_CODE = re.compile(r"\[(RED|BLACK|BLUE|GREEN|CYAN|MAGENTA|WHITE|YELLOW|COLOR\d+)\]", re.I)
COUNT = "=SUMPRODUCT(--(ABS("
SCRATCH = "_layout_lint"
VENDOR_TABS = ("RFS_info", "RFS_fvalues", "RFS_cashflow")  # their layout is the vendor's
NUMBER, CURRENCY, PERCENT, TEXT = 16, 8, 128, 256  # com.sun.star.util.NumberFormat
ALL_FLAGS = 1023  # com.sun.star.sheet.CellFlags: everything


def col(n):
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def addr(sheet, c, r):
    return "%s!%s%d" % (sheet.Name, col(c), r + 1)


def used(sheet):
    cur = sheet.createCursor()
    cur.gotoEndOfUsedArea(False)
    a = cur.RangeAddress
    return a.EndRow + 1, a.EndColumn + 1


def format_ranges(sheet, ranges):
    """The uniformly formatted parts of a set of ranges (XSheetCellRanges)."""
    for a in ranges.getRangeAddresses():
        rng = sheet.getCellRangeByPosition(a.StartColumn, a.StartRow, a.EndColumn, a.EndRow)
        yield from rng.getCellFormatRanges()


def justify(cell):
    return cell.HoriJustify.value  # STANDARD, LEFT, CENTER, RIGHT, BLOCK, REPEAT


class Lint:
    def __init__(self, doc):
        self.doc = doc
        self.problems = []
        self.cells_checked = 0

    def run(self):
        sheets = self.doc.Sheets
        tabs = [s for s in sheets]
        for sheet in tabs:
            if not sheet.isProtected():
                self.problems.append("%s: the tab is not protected" % sheet.Name)
            else:
                sheet.unprotect("")
        self.open_scratch()
        self.alarm_styles = self.find_alarm_styles()
        for sheet in tabs:
            self.protection(sheet)
            self.conditional_formats(sheet)
            self.negative_zero(sheet)
            self.fit(sheet)
        self.close_scratch()
        return self.problems

    # -- protection -----------------------------------------------------------------------

    def protection(self, sheet):
        rows, cols = used(sheet)
        area = sheet.getCellRangeByPosition(0, 0, cols - 1, rows - 1)
        # A yellow input cell may hold the user's own formula; any other unlocked cell is wrong.
        for rng in area.getCellFormatRanges():
            if not rng.CellProtection.IsLocked and rng.CellBackColor != INPUT_FILL:
                self.problems.append("%s: unlocked but not a yellow input cell" % rng.AbsoluteName)

    # -- display ------------------------------------------------------------------------

    def negative_zero(self, sheet):
        """A rounding remainder shown as "-0" (design 17.2)."""
        if sheet.Name in VENDOR_TABS:
            return
        rows, cols = used(sheet)
        data = sheet.getCellRangeByPosition(0, 0, cols - 1, rows - 1).getDataArray()
        for r, row in enumerate(data):
            for c, v in enumerate(row):
                if isinstance(v, float) and -1 < v < 0:
                    shown = sheet.getCellByPosition(c, r).getString()
                    if re.fullmatch(r"-0([.,]0*)?%?", shown.strip()):
                        self.problems.append("%s: shows %r" % (addr(sheet, c, r), shown))

    # -- conditional formats --------------------------------------------------------------

    def find_alarm_styles(self):
        styles = self.doc.StyleFamilies.getByName("CellStyles")
        return {n for n in styles.ElementNames
                if styles.getByName(n).CellBackColor == ALARM_FILL}

    def conditional_formats(self, sheet):
        styles = self.doc.StyleFamilies.getByName("CellStyles")
        formats = self.doc.NumberFormats
        alarm_cells = []
        for cf in sheet.ConditionalFormats.ConditionalFormats:
            ranges = cf.Range
            names = []
            for i in range(cf.Count):
                entry = cf.getByIndex(i)
                if hasattr(entry, "StyleName"):
                    names.append(entry.StyleName)
            where = ranges.RangeAddressesAsString
            content = ranges.queryContentCells(1 | 2 | 4 | 16)  # values, dates, strings, formulas
            if not content.getCount():
                self.problems.append("%s: conditional format on empty cells only" % where)
            coloured = [n for n in names if styles.hasByName(n)
                        and styles.getByName(n).CharColor not in (-1,)
                        and not styles.getByName(n).getPropertyValue("NumberFormat")]
            if any(n in self.alarm_styles for n in names):
                alarm_cells.append(ranges)
            if not coloured:
                continue
            for rng in format_ranges(sheet, ranges):
                code = formats.getByKey(rng.NumberFormat).FormatString
                if COLOUR_CODE.search(code):
                    self.problems.append(
                        "%s: number format %r overrides the font colour of conditional style %s"
                        % (rng.AbsoluteName, code, coloured[0]))
        rows, cols = used(sheet)
        data = sheet.getCellRangeByPosition(0, 0, cols - 1, rows - 1).getFormulaArray()
        for r, row in enumerate(data):
            for c, f in enumerate(row):
                if f.startswith(COUNT):
                    cell = sheet.getCellByPosition(c, r)
                    if not any(rngs.queryIntersection(cell.RangeAddress).getCount()
                               for rngs in alarm_cells):
                        self.problems.append("%s: a discrepancy count without the Alarm format"
                                             % addr(sheet, c, r))

    # -- fit ------------------------------------------------------------------------------

    def fit(self, sheet):
        if sheet.Name in VENDOR_TABS or sheet.Name == SCRATCH:
            return
        rows, cols = used(sheet)
        area = sheet.getCellRangeByPosition(0, 0, cols - 1, rows - 1)
        data = area.getDataArray()
        formulas = area.getFormulaArray()
        col_vis = [sheet.Columns.getByIndex(c).IsVisible for c in range(cols + 20)]
        widths = [sheet.Columns.getByIndex(c).Width for c in range(cols + 20)]
        row_objs = [sheet.Rows.getByIndex(r) for r in range(rows)]
        row_vis = [x.IsVisible for x in row_objs]
        heights = [x.Height for x in row_objs]
        automatic = [x.OptimalHeight for x in row_objs]

        def empty(r, c):
            return c >= cols or (data[r][c] == "" and not formulas[r][c])

        # Input rows keep an automatic height, so that they grow as the user types (design 17.3).
        inputs = set()
        for rng in area.getCellFormatRanges():
            if not rng.CellProtection.IsLocked:
                a = rng.RangeAddress
                inputs.update((r, c) for r in range(a.StartRow, a.EndRow + 1)
                              for c in range(a.StartColumn, a.EndColumn + 1))
        for r in sorted({r for r, _ in inputs}):
            if r < rows and row_vis[r] and not automatic[r]:
                self.problems.append("%s!%d: an input row with a fixed height" % (sheet.Name, r + 1))

        # Wrapped text in a hidden column still sets an automatic row's height: blank lines.
        for c in range(cols):
            if col_vis[c]:
                continue
            for r in range(rows):
                if row_vis[r] and automatic[r] and isinstance(data[r][c], str) and data[r][c] \
                        and sheet.getCellByPosition(c, r).IsTextWrapped:
                    self.problems.append("%s: wrapped text in a hidden column sets the row's "
                                         "height" % addr(sheet, c, r))

        # Numbers: in each column, the largest positive and the most negative value stand for the
        # rest (the column shares a format). Hidden columns too: a folded group can be unfolded.
        numbers = {}
        for r in range(rows):
            if not row_vis[r]:
                continue
            for c in range(cols):
                v = data[r][c]
                if isinstance(v, float):
                    key = (c, v < 0)
                    if key not in numbers or abs(v) > abs(data[numbers[key]][c]):
                        numbers[key] = r
        candidates = {(r, c) for (c, _), r in numbers.items()}
        for r in range(rows):
            if not row_vis[r]:
                continue
            for c in range(cols):
                if col_vis[c] and isinstance(data[r][c], str) and data[r][c] != "":
                    candidates.add((r, c))

        for r, c in sorted(candidates):
            cell = sheet.getCellByPosition(c, r)
            if cell.getError() or cell.ShrinkToFit or cell.RotateAngle:
                continue
            cur = sheet.createCursorByRange(cell)
            cur.collapseToMergedArea()
            m = cur.RangeAddress
            merged = (m.EndColumn, m.EndRow) != (c, r)
            if merged and (m.StartColumn, m.StartRow) != (c, r):
                continue  # a hidden part of a merged area
            w = sum(widths[k] for k in range(m.StartColumn, m.EndColumn + 1) if col_vis[k])
            if not col_vis[c]:
                if not isinstance(data[r][c], float):
                    continue
                w = widths[c]  # a number in a folded column, measured as if unfolded
            if self.in_filter_header(cell):
                w -= FILTER_BUTTON  # the drop-down button covers the cell's right side
            h = sum(heights[k] for k in range(m.StartRow, m.EndRow + 1) if row_vis[k])
            fixed = not all(automatic[k] for k in range(m.StartRow, m.EndRow + 1))
            value = data[r][c]
            text = value.rstrip() if isinstance(value, str) else value
            where = addr(sheet, c, r)
            show = repr(str(text)[:40])
            self.cells_checked += 1
            self.load(cell, formulas[r][c], text)
            wrapped = cell.IsTextWrapped
            if wrapped:
                # The program fits an automatic row itself; a fixed height needs the margin.
                need_h = self.height(int(w / WIDER) if fixed else w)
                if need_h > h + SLACK:
                    self.problems.append("%s: needs a height of %d, has %d (%s)"
                                         % (where, need_h, h, show))
                if isinstance(text, str):
                    # Prose in automatic rows (notes, descriptions) may break at its hyphens
                    # ("GL 9220-4012"); a heading in a fixed row may not ("Late-/year").
                    self.breaks(where, text, w, show, hyphens=fixed)
                continue
            need_w = self.width()
            if fixed:
                need_h = self.height(None)
                if need_h > h + SLACK:
                    self.problems.append("%s: needs a height of %d, has %d (%s)"
                                         % (where, need_h, h, show))
            if need_w * WIDER <= w + SLACK:
                continue
            if isinstance(text, float) or merged:
                self.problems.append("%s: needs a width of %d, has %d (%s)"
                                     % (where, int(need_w * WIDER), w, show))
                continue
            if (r, c) in inputs or self.numeric_format(cell):
                continue  # text the user typed into an input cell or a number column (17.3)
            # Text overflows into empty neighbours, on the side its alignment points to.
            side = justify(cell)
            room = w
            if side in ("STANDARD", "LEFT", "CENTER"):
                k = c + 1
                while room < need_w * WIDER and empty(r, k):
                    room += widths[k] if col_vis[k] else 0
                    k += 1
            if side in ("RIGHT", "CENTER"):
                k = c - 1
                while room < need_w * WIDER and k >= 0 and empty(r, k):
                    room += widths[k] if col_vis[k] else 0
                    k -= 1
            if room + SLACK < need_w * WIDER:
                self.problems.append("%s: needs a width of %d, has %d with its empty neighbours "
                                     "(%s)" % (where, int(need_w * WIDER), room, show))

    def breaks(self, where, text, w, show, hyphens=True):
        """Wrapped text must not break inside a word, at a hyphen ("Late-/year") or anywhere
        else, even when it renders `WIDER` wider. Lines are filled as the programs fill them:
        as many pieces as fit, a piece being a word or the part of a word up to a hyphen."""
        for avail in (w, w / WIDER):
            for paragraph in text.split("\n"):
                pieces = []  # (text, joined to the previous piece without a space)
                for word in paragraph.split():
                    parts = re.findall(r"[^-]+-?|-", word)
                    pieces += [(p, i > 0) for i, p in enumerate(parts)]
                line = ""
                for piece, joined in pieces:
                    trial = line + ("" if joined or not line else " ") + piece
                    if self.text_width(trial) <= avail:
                        line = trial
                        continue
                    if not line or self.text_width(piece) > avail:
                        self.problems.append("%s: a word is wider than the cell (%s)"
                                             % (where, show))
                        return
                    if joined and hyphens:
                        self.problems.append("%s: breaks inside a word, after %r (%s)"
                                             % (where, line.split()[-1], show))
                        return
                    line = piece

    def numeric_format(self, cell):
        """A number format other than General: the cell is meant for a number."""
        fmt = self.doc.NumberFormats.getByKey(cell.NumberFormat)
        if fmt.FormatString.lower().startswith("general") or fmt.Type & TEXT:
            return False
        return bool(fmt.Type & (NUMBER | CURRENCY | PERCENT))

    # -- measuring on the scratch tab -----------------------------------------------------

    def open_scratch(self):
        sheets = self.doc.Sheets
        sheets.insertNewByName(SCRATCH, sheets.Count)
        self.scratch = sheets.getByName(SCRATCH)

    def close_scratch(self):
        self.doc.Sheets.removeByName(SCRATCH)

    def in_filter_header(self, cell):
        """Whether the cell is in an AutoFilter's header row, under a drop-down button."""
        a = cell.CellAddress
        ranges = self.doc.getPropertyValue("UnnamedDatabaseRanges")
        if not ranges.hasByTable(a.Sheet):
            return False
        db = ranges.getByTable(a.Sheet)
        area = db.DataArea
        return bool(db.AutoFilter) and a.Row == area.StartRow \
            and area.StartColumn <= a.Column <= area.EndColumn

    def load_cell(self, cell):
        value = cell.getValue() if cell.getType().value == "VALUE" else cell.getString().rstrip()
        self.load(cell, cell.getFormula(), value)
        return value

    def needed_width(self, cell):
        """The width a cell needs with `WIDER`'s margin: its content on one line, or, wrapped,
        its widest word."""
        value = self.load_cell(cell)
        if cell.IsTextWrapped and isinstance(value, str):
            pieces = [p for word in value.split() for p in re.findall(r"[^-]+-?|-", word)]
            need = max([self.text_width(p) for p in pieces] or [0])
        else:
            need = self.width()
        button = FILTER_BUTTON if self.in_filter_header(cell) else 0
        return int(need * WIDER) + SLACK + button

    def needed_height(self, cell, width):
        """The height a wrapped cell needs at `width` less `WIDER`'s margin (and less a filter
        button's width, under one)."""
        if self.in_filter_header(cell):
            width -= FILTER_BUTTON
        self.load_cell(cell)
        return self.height(int(width / WIDER)) + SLACK

    def load(self, cell, formula, text):
        """Copies the cell, with its formats, to the scratch tab's A1, holding `text`."""
        s = self.scratch
        target = s.getCellByPosition(0, 0)
        s.getCellRangeByPosition(0, 0, 0, 0).clearContents(ALL_FLAGS)
        dest = CellAddress()
        dest.Sheet, dest.Column, dest.Row = target.CellAddress.Sheet, 0, 0
        s.copyRange(dest, cell.RangeAddress)
        if isinstance(text, float):
            if formula.startswith("="):
                target.setValue(text)
        elif formula.startswith("=") or text != cell.getString():
            target.setString(text)  # a formula's result, or text without trailing spaces
        self.target = target
        self.widths = {}

    def width(self):
        column = self.scratch.Columns.getByIndex(0)
        wrapped = self.target.IsTextWrapped
        self.target.IsTextWrapped = False
        column.OptimalWidth = True
        self.target.IsTextWrapped = wrapped
        return column.Width

    def height(self, width):
        if width is not None:
            self.scratch.Columns.getByIndex(0).Width = int(width)
        row = self.scratch.Rows.getByIndex(0)
        row.OptimalHeight = True
        return row.Height

    def text_width(self, text):
        """The width `text` needs in the loaded cell's format, on one line."""
        if text not in self.widths:
            probe = self.scratch.getCellByPosition(1, 0)
            dest = CellAddress()
            dest.Sheet, dest.Column, dest.Row = probe.CellAddress.Sheet, 1, 0
            self.scratch.copyRange(dest, self.target.RangeAddress)
            probe.IsTextWrapped = False
            probe.setString(text)
            column = self.scratch.Columns.getByIndex(1)
            column.OptimalWidth = True
            self.widths[text] = column.Width
            self.scratch.getCellRangeByPosition(1, 0, 1, 0).clearContents(ALL_FLAGS)
        return self.widths[text]


def lint(doc):
    return Lint(doc).run()


def main():
    path = sys.argv[1]
    with Session() as s:
        doc = s.open_copy(path)
        doc.calculateAll()
        linter = Lint(doc)
        problems = linter.run()
        doc.close(True)
    print("%d cells measured" % linter.cells_checked)
    for p in problems:
        print("FAIL " + p)
    print("%d problems" % len(problems))
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
