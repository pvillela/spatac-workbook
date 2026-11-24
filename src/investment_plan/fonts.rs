//! Each cell's font, read from the file itself.
//!
//! LibreOffice writes `applyFont="false"` on a cell format whose font is inherited from its cell
//! style, while still naming the font in `fontId`. umya-spreadsheet drops the font of such a cell,
//! and on writing gives it the new workbook's default (Calibri 11), so the Investment Plan came out
//! in a different font from its source (design 14.2). The rule here: a cell's font is the one its
//! cell format names, whatever `applyFont` says.

use crate::diff::protection::{attrs, formats_in, sheet_parts};
use regex::Regex;
use std::collections::HashMap;
use std::io::Read;
use std::path::Path;
use std::sync::LazyLock;
use umya_spreadsheet::Style;

/// A font as the plan needs it.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct FontSpec {
    pub name: Option<String>,
    pub size: Option<f64>,
    pub bold: bool,
    pub italic: bool,
    /// An `rgb` colour, e.g. `FFC00000`; theme colours are left out.
    pub color: Option<String>,
}

impl FontSpec {
    /// Gives `style` this font.
    pub fn apply(&self, style: &mut Style) {
        let font = style.font_mut();
        if let Some(name) = &self.name {
            font.set_name(name.clone());
        }
        if let Some(size) = self.size {
            font.set_size(size);
        }
        font.set_bold(self.bold);
        font.set_italic(self.italic);
        if let Some(color) = &self.color {
            font.color_mut().set_argb_str(color);
        }
    }
}

/// The fonts of a workbook's cells.
pub struct CellFonts {
    /// The default font: that of cell format 0.
    pub default: FontSpec,
    /// Each tab's cells, as (row, column), 1-based, keyed by tab name.
    pub cells: HashMap<String, HashMap<(u32, u32), FontSpec>>,
}

/// Reads the font of every cell of the tabs in `sheets`.
pub fn cell_fonts(path: &Path, sheets: &[&str]) -> Result<CellFonts, String> {
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
    let styles = part("xl/styles.xml").unwrap_or_default();
    let fonts = fonts_in(&styles);
    let format_fonts: Vec<FontSpec> = formats_in(&styles, "cellXfs")
        .iter()
        .map(|f| {
            f.font
                .and_then(|i| fonts.get(i))
                .cloned()
                .unwrap_or_default()
        })
        .collect();
    let default = format_fonts.first().cloned().unwrap_or_default();
    let mut cells = HashMap::new();
    for (name, target) in sheet_parts(&book, &rels) {
        if !sheets.contains(&name.as_str()) {
            continue;
        }
        let xml = part(&target).unwrap_or_default();
        cells.insert(name, cells_in(&xml, &format_fonts));
    }
    Ok(CellFonts { default, cells })
}

/// The fonts of `<fonts>`, in order.
fn fonts_in(styles: &str) -> Vec<FontSpec> {
    static FONTS: LazyLock<Regex> = LazyLock::new(|| {
        Regex::new(r"(?s)<(?:\w+:)?fonts\b[^>]*>(.*?)</(?:\w+:)?fonts>").expect("valid pattern")
    });
    static FONT: LazyLock<Regex> = LazyLock::new(|| {
        Regex::new(r"(?s)<(?:\w+:)?font\b[^>]*?(?:/>|>(.*?)</(?:\w+:)?font>)")
            .expect("valid pattern")
    });
    static CHILD: LazyLock<Regex> =
        LazyLock::new(|| Regex::new(r"<(?:\w+:)?(\w+)\b[^>]*>").expect("valid pattern"));
    let Some(body) = FONTS.captures(styles).and_then(|c| c.get(1)) else {
        return Vec::new();
    };
    FONT.captures_iter(body.as_str())
        .map(|c| {
            let mut spec = FontSpec::default();
            let inner = c.get(1).map_or("", |m| m.as_str());
            for child in CHILD.captures_iter(inner) {
                let a = attrs(child.get(0).unwrap().as_str());
                let on = a.get("val").is_none_or(|v| !matches!(*v, "0" | "false"));
                match &child[1] {
                    "name" => spec.name = a.get("val").map(|v| (*v).to_owned()),
                    "sz" => spec.size = a.get("val").and_then(|v| v.parse().ok()),
                    "b" => spec.bold = on,
                    "i" => spec.italic = on,
                    "color" => spec.color = a.get("rgb").map(|v| (*v).to_owned()),
                    _ => {}
                }
            }
            spec
        })
        .collect()
}

/// The font of every cell in a sheet part, by its cell format.
fn cells_in(sheet: &str, format_fonts: &[FontSpec]) -> HashMap<(u32, u32), FontSpec> {
    static CELL: LazyLock<Regex> =
        LazyLock::new(|| Regex::new(r"<(?:\w+:)?c\s[^>]*>").expect("valid pattern"));
    static REF: LazyLock<Regex> =
        LazyLock::new(|| Regex::new(r"^([A-Z]+)([0-9]+)$").expect("valid pattern"));
    CELL.find_iter(sheet)
        .filter_map(|m| {
            let a = attrs(m.as_str());
            let s: usize = a.get("s").and_then(|v| v.parse().ok()).unwrap_or(0);
            let caps = REF.captures(a.get("r")?)?;
            let col = caps[1]
                .bytes()
                .fold(0u32, |n, b| n * 26 + u32::from(b - b'A' + 1));
            Some(((caps[2].parse().ok()?, col), format_fonts.get(s)?.clone()))
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    const STYLES: &str = r#"<styleSheet><fonts count="3">
        <font><sz val="10"/><name val="Arial"/><family val="2"/></font>
        <font><b val="true"/><sz val="10"/><name val="Arial"/></font>
        <font><b/><i val="false"/><sz val="14"/><color rgb="FFC00000"/><name val="Arial"/></font>
        </fonts><cellStyleXfs count="1"><xf fontId="1"/></cellStyleXfs>
        <cellXfs count="3"><xf fontId="0" xfId="0"/><xf fontId="1" xfId="0" applyFont="false"/>
        <xf fontId="2" xfId="0" applyFont="true"/></cellXfs></styleSheet>"#;

    #[test]
    fn fonts_are_read_in_order() {
        let fonts = fonts_in(STYLES);
        assert_eq!(fonts.len(), 3);
        assert_eq!(fonts[0].name.as_deref(), Some("Arial"));
        assert_eq!(fonts[0].size, Some(10.0));
        assert!(!fonts[0].bold);
        assert!(fonts[1].bold);
        assert!(fonts[2].bold && !fonts[2].italic);
        assert_eq!(fonts[2].color.as_deref(), Some("FFC00000"));
    }

    #[test]
    fn a_cells_font_is_its_formats_whatever_apply_says() {
        let fonts = fonts_in(STYLES);
        let formats: Vec<FontSpec> = formats_in(STYLES, "cellXfs")
            .iter()
            .map(|f| fonts[f.font.unwrap()].clone())
            .collect();
        let sheet =
            r#"<row r="3"><c r="A3"/><c r="B3" s="1"/><c r="C3" s="2" t="n"><v>1</v></c></row>"#;
        let cells = cells_in(sheet, &formats);
        assert!(!cells[&(3, 1)].bold);
        assert!(cells[&(3, 2)].bold);
        assert_eq!(cells[&(3, 3)].size, Some(14.0));
    }
}
