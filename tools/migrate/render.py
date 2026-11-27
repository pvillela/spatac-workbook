"""Renders parts of a workbook to PDF, to look at a build before handing it over (design-history.md
section 17.2).

Usage:

    /opt/libreoffice26.2/program/python tools/migrate/render.py WORKBOOK.xlsx OUT.pdf \
        [--scale PERCENT] [TAB[!RANGE] ...]

With no TAB, every tab's used area is rendered. Each range prints at SCALE percent (default 100)
on landscape A3 pages, with row and column headings, so that what the page shows is what the
screen shows at that zoom. The workbook is opened read-only, as a copy.
"""

import argparse
import os
import sys

import uno  # noqa: F401
from com.sun.star.table import CellRangeAddress

from uno_session import Session, prop


def used_address(sheet):
    cur = sheet.createCursor()
    cur.gotoEndOfUsedArea(False)
    a = cur.RangeAddress
    out = CellRangeAddress()
    out.Sheet, out.StartColumn, out.StartRow = a.Sheet, 0, 0
    out.EndColumn, out.EndRow = a.EndColumn, a.EndRow
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("workbook")
    ap.add_argument("output")
    ap.add_argument("--scale", type=int, default=100)
    ap.add_argument("parts", nargs="*")
    args = ap.parse_args()
    if os.path.exists(args.output):
        raise SystemExit("%s exists: the script never overwrites a file" % args.output)
    with Session() as s:
        doc = s.open_copy(args.workbook)
        doc.calculateAll()
        wanted = {}
        for part in args.parts or [sh.Name for sh in doc.Sheets]:
            tab, _, rng = part.partition("!")
            wanted[tab] = rng
        styles = doc.StyleFamilies.getByName("PageStyles")
        for i, sheet in enumerate(doc.Sheets):
            name = "Render%d" % i
            style = doc.createInstance("com.sun.star.style.PageStyle")
            styles.insertByName(name, style)
            style.IsLandscape = True
            style.Width, style.Height = 42000, 29700  # A3 landscape, 1/100 mm
            for side in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
                setattr(style, side, 800)
            style.PageScale = args.scale
            style.PrintHeaders = True
            style.PrintGrid = False
            style.HeaderIsOn = True
            style.FooterIsOn = False
            right = style.RightPageHeaderContent
            right.CenterText.setString(sheet.Name)
            style.RightPageHeaderContent = right
            sheet.PageStyle = name
            if sheet.Name not in wanted:
                continue
            if wanted[sheet.Name]:
                area = sheet.getCellRangeByName(wanted[sheet.Name]).RangeAddress
            else:
                area = used_address(sheet)
            sheet.setPrintAreas((area,))
            sheet.setPrintTitleRows(False)
            sheet.setPrintTitleColumns(False)
        # Only the tabs asked for: the PDF export leaves hidden tabs out.
        doc.CurrentController.setActiveSheet(doc.Sheets.getByName(next(iter(wanted))))
        for sh in doc.Sheets:
            if sh.Name not in wanted:
                sh.IsVisible = False
        url = uno.systemPathToFileUrl(os.path.abspath(args.output))
        doc.storeToURL(url, (prop("FilterName", "calc_pdf_Export"),))
        doc.close(True)
    print("Wrote %s" % args.output)


if __name__ == "__main__":
    sys.exit(main())
