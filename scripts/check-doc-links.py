#!/usr/bin/env python3
"""Checks the links between pages of the mdBook in docs/.

Every relative link must point at a page that exists, and every #anchor at a heading on that
page (anchors as mdBook makes them: lower case, punctuation dropped, spaces as hyphens).
Exits 1 if any link is broken. Run from anywhere: paths are relative to this script.

    python3 scripts/check-doc-links.py
"""

import pathlib
import re
import sys

DOCS = pathlib.Path(__file__).resolve().parent.parent / "docs"
LINK = re.compile(r"\[[^\]]*\]\(([^)\s]+)\)")
HEADING = re.compile(r"^#{1,6}\s+(.*?)\s*$", re.M)


def slug(text):
    text = re.sub(r"`", "", text).strip().lower()
    text = re.sub(r"[^\w\s-]", "", text)
    return re.sub(r"\s", "-", text)


def anchors(path):
    body = re.sub(r"```.*?```", "", path.read_text(encoding="utf-8"), flags=re.S)
    return {slug(h) for h in HEADING.findall(body)}


def main():
    pages = sorted(p for p in DOCS.rglob("*.md") if "archive" not in p.parts)
    broken = []
    for page in pages:
        body = re.sub(r"```.*?```", "", page.read_text(encoding="utf-8"), flags=re.S)
        for target in LINK.findall(body):
            if re.match(r"^[a-z]+:", target):
                continue
            file, _, anchor = target.partition("#")
            dest = page if not file else (page.parent / file).resolve()
            if not dest.exists():
                broken.append("%s: %s (no such page)" % (page.relative_to(DOCS), target))
            elif anchor and anchor not in anchors(dest):
                broken.append("%s: %s (no such heading)" % (page.relative_to(DOCS), target))
    for b in broken:
        print(b)
    print("%d pages, %d broken links" % (len(pages), len(broken)))
    return 1 if broken else 0


if __name__ == "__main__":
    sys.exit(main())
