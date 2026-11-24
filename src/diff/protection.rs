//! Which cells are unlocked, read from the file itself.
//!
//! A cell's format (`cellXfs`) may carry its own `<protection>` element, or leave protection to
//! its cell style (`cellStyleXfs`). Applications disagree on the `applyProtection` attribute that
//! goes with it: LibreOffice writes `applyProtection="false"` on a cell format whose
//! `<protection locked="false"/>` is the cell's real state, inherited from its style. Excel and
//! LibreOffice both treat such a cell as unlocked. umya-spreadsheet instead drops protection
//! altogether when `applyProtection` is false, so the same unlocked cell reads as locked in one
//! file and unlocked in another, and the diff reported cells as newly unlocked when nothing had
//! changed (design 13.8).
//!
//! The rule here: a cell format's own `<protection>` element decides; without one, its cell
//! style's does; without either, the cell is locked, the default.

use regex::Regex;
use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::io::Read;
use std::path::Path;
use std::sync::LazyLock;

/// The unlocked cells of each tab, as (row, column), 1-based, keyed by tab name.
pub type Unlocked = HashMap<String, BTreeSet<(u32, u32)>>;

/// Reads the unlocked cells of every tab of the `.xlsx` file at `path`.
pub fn unlocked_cells(path: &Path) -> Result<Unlocked, String> {
    let file =
        std::fs::File::open(path).map_err(|e| format!("cannot read {}: {e}", path.display()))?;
    let mut zip =
        zip::ZipArchive::new(file).map_err(|e| format!("cannot read {}: {e}", path.display()))?;
    let mut part = |name: &str| -> Option<String> {
        let mut text = String::new();
        zip.by_name(name).ok()?.read_to_string(&mut text).ok()?;
        Some(text)
    };
    let (Some(book), Some(rels)) = (part("xl/workbook.xml"), part("xl/_rels/workbook.xml.rels"))
    else {
        return Err(format!("{} is not an .xlsx workbook", path.display()));
    };
    let unlocked_xfs = part("xl/styles.xml").map_or_else(BTreeSet::new, |s| unlocked_formats(&s));
    let mut out = Unlocked::new();
    for (name, target) in sheet_parts(&book, &rels) {
        let cells =
            part(&target).map_or_else(BTreeSet::new, |xml| cells_with_formats(&xml, &unlocked_xfs));
        out.insert(name, cells);
    }
    Ok(out)
}

static ATTR: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r#"([A-Za-z:]+)\s*=\s*"([^"]*)""#).expect("valid pattern"));

pub(crate) fn attrs(tag: &str) -> BTreeMap<&str, &str> {
    ATTR.captures_iter(tag)
        .map(|c| (c.get(1).unwrap().as_str(), c.get(2).unwrap().as_str()))
        .collect()
}

/// Each tab's name and the zip path of its part, in workbook order.
pub(crate) fn sheet_parts(book: &str, rels: &str) -> Vec<(String, String)> {
    static SHEET: LazyLock<Regex> =
        LazyLock::new(|| Regex::new(r"<(?:\w+:)?sheet\s[^>]*>").expect("valid pattern"));
    static REL: LazyLock<Regex> =
        LazyLock::new(|| Regex::new(r"<(?:\w+:)?Relationship\s[^>]*>").expect("valid pattern"));
    let targets: HashMap<&str, &str> = REL
        .find_iter(rels)
        .filter_map(|m| {
            let a = attrs(m.as_str());
            Some((*a.get("Id")?, *a.get("Target")?))
        })
        .collect();
    SHEET
        .find_iter(book)
        .filter_map(|m| {
            let a = attrs(m.as_str());
            let target = targets.get(a.get("r:id")?)?;
            let path = match target.strip_prefix('/') {
                Some(absolute) => absolute.to_owned(),
                None => format!("xl/{target}"),
            };
            Some((unescape(a.get("name")?), path))
        })
        .collect()
}

pub(crate) fn unescape(s: &str) -> String {
    s.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&amp;", "&")
}

/// The indexes of the cell formats (`cellXfs`) whose cells are unlocked.
fn unlocked_formats(styles: &str) -> BTreeSet<usize> {
    let style_locks = formats_in(styles, "cellStyleXfs");
    formats_in(styles, "cellXfs")
        .iter()
        .enumerate()
        .filter(|(_, f)| {
            let inherited = f
                .style
                .and_then(|i| style_locks.get(i))
                .and_then(|s| s.locked);
            !f.locked.or(inherited).unwrap_or(true)
        })
        .map(|(i, _)| i)
        .collect()
}

pub(crate) struct Format {
    /// The `locked` value of the format's own `<protection>`, if it has one.
    locked: Option<bool>,
    /// The cell style it is based on (`xfId`).
    style: Option<usize>,
    /// Its font (`fontId`).
    pub(crate) font: Option<usize>,
}

pub(crate) fn formats_in(styles: &str, list: &str) -> Vec<Format> {
    static XF: LazyLock<Regex> = LazyLock::new(|| {
        Regex::new(r"(?s)<(?:\w+:)?xf\b[^>]*?(?:/>|>.*?</(?:\w+:)?xf>)").expect("valid pattern")
    });
    static PROTECTION: LazyLock<Regex> =
        LazyLock::new(|| Regex::new(r"<(?:\w+:)?protection\b[^>]*>").expect("valid pattern"));
    let open = Regex::new(&format!(
        r"(?s)<(?:\w+:)?{list}\b[^>]*>(.*?)</(?:\w+:)?{list}>"
    ))
    .expect("valid pattern");
    let Some(body) = open.captures(styles).and_then(|c| c.get(1)) else {
        return Vec::new();
    };
    XF.find_iter(body.as_str())
        .map(|m| {
            let xf = m.as_str();
            let head = &xf[..xf.find('>').unwrap_or(xf.len())];
            let locked = PROTECTION.find(xf).map(|p| {
                attrs(p.as_str())
                    .get("locked")
                    .is_none_or(|v| !matches!(*v, "0" | "false"))
            });
            let head = attrs(head);
            let style = head.get("xfId").and_then(|v| v.parse().ok());
            let font = head.get("fontId").and_then(|v| v.parse().ok());
            Format {
                locked,
                style,
                font,
            }
        })
        .collect()
}

/// The (row, column) of every cell in a sheet part whose format is unlocked.
fn cells_with_formats(sheet: &str, unlocked_xfs: &BTreeSet<usize>) -> BTreeSet<(u32, u32)> {
    static CELL: LazyLock<Regex> =
        LazyLock::new(|| Regex::new(r"<(?:\w+:)?c\s[^>]*>").expect("valid pattern"));
    static REF: LazyLock<Regex> =
        LazyLock::new(|| Regex::new(r"^([A-Z]+)([0-9]+)$").expect("valid pattern"));
    CELL.find_iter(sheet)
        .filter_map(|m| {
            let a = attrs(m.as_str());
            let s: usize = a.get("s").and_then(|v| v.parse().ok()).unwrap_or(0);
            if !unlocked_xfs.contains(&s) {
                return None;
            }
            let caps = REF.captures(a.get("r")?)?;
            let col = caps[1]
                .bytes()
                .fold(0u32, |n, b| n * 26 + u32::from(b - b'A' + 1));
            Some((caps[2].parse().ok()?, col))
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    // The yellow input cell B9 as two applications wrote it: LibreOffice's own build, with
    // applyProtection="false", and the same workbook saved after editing, with "true".
    const BUILT: &str = r#"<styleSheet><cellStyleXfs count="2"><xf numFmtId="0"/>
        <xf numFmtId="164" applyProtection="true"><protection locked="false" hidden="false"/></xf>
        </cellStyleXfs><cellXfs count="2"><xf numFmtId="0" xfId="0"/>
        <xf numFmtId="168" xfId="1" applyProtection="false"><alignment vertical="bottom"/>
        <protection locked="false" hidden="false"/></xf></cellXfs></styleSheet>"#;
    const RESAVED: &str = r#"<styleSheet><cellStyleXfs count="2"><xf numFmtId="0"/>
        <xf numFmtId="164" applyProtection="true"><protection locked="false"/></xf>
        </cellStyleXfs><cellXfs count="2"><xf numFmtId="0" xfId="0"/>
        <xf numFmtId="168" xfId="1" applyProtection="true"><protection locked="false"/></xf>
        </cellXfs></styleSheet>"#;

    #[test]
    fn a_cell_formats_own_protection_decides_whatever_apply_says() {
        assert_eq!(unlocked_formats(BUILT), BTreeSet::from([1]));
        assert_eq!(unlocked_formats(RESAVED), BTreeSet::from([1]));
    }

    #[test]
    fn without_its_own_protection_a_format_inherits_its_styles() {
        let styles = r#"<cellStyleXfs><xf/><xf><protection locked="0"/></xf></cellStyleXfs>
            <cellXfs><xf xfId="1"/><xf xfId="0"/><xf xfId="1"><protection locked="1"/></xf></cellXfs>"#;
        assert_eq!(unlocked_formats(styles), BTreeSet::from([0]));
    }

    #[test]
    fn cells_are_found_by_their_format() {
        let sheet = r#"<sheetData><row r="9"><c r="A9" s="0"/><c r="B9" s="1" t="n"><v>5</v></c>
            <c r="AA10" s="1"/></row></sheetData>"#;
        assert_eq!(
            cells_with_formats(sheet, &BTreeSet::from([1])),
            BTreeSet::from([(9, 2), (10, 27)])
        );
    }

    #[test]
    fn sheet_parts_follow_the_relationships() {
        let book = r#"<sheets><sheet name="A &amp; B" sheetId="1" r:id="rId3"/><sheet name="C" sheetId="2" r:id="rId4"/></sheets>"#;
        let rels = r#"<Relationships><Relationship Id="rId3" Target="worksheets/sheet1.xml"/><Relationship Id="rId4" Target="/xl/worksheets/sheet2.xml"/></Relationships>"#;
        assert_eq!(
            sheet_parts(book, rels),
            [
                ("A & B".to_owned(), "xl/worksheets/sheet1.xml".to_owned()),
                ("C".to_owned(), "xl/worksheets/sheet2.xml".to_owned())
            ]
        );
    }
}
