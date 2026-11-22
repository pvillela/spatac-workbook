//! Investment Plan extraction.
//!
//! [`extract`] reads a Spatac workbook and writes the Investment Plan as a workbook of its own,
//! for the Board, the Treasurer and the broker (Reserve Fund Investment Policy s.3.2):
//!
//! - **Cover**: what the plan is, the source file, when it was generated, the base year, lines
//!   to record the meeting that approved it, and where the reasons for changes from the RFS are
//!   recorded.
//! - **Medium_Term_Cashflow**: the cash section and the RFS-versus-Revised comparison. The
//!   investment-balance comparison and notes are dropped.
//! - **Schedule_of_Investments**: held and planned investments with their cash flows by year.
//!
//! Formulas that refer only to these two tabs stay formulas. A formula that refers to any other
//! tab, directly or through a defined name, becomes its cached value. Every cell is locked and
//! every sheet protected without a password.
//!
//! Content is located only through defined names (design section 3.4), so the extraction keeps
//! working when rows or columns move.

use crate::diff::formula::{self, Token};
use crate::diff::model::parse_range;
use regex::Regex;
use std::collections::{BTreeMap, BTreeSet};
use std::path::{Path, PathBuf};
use std::sync::LazyLock;
use umya_spreadsheet::structs::DefinedName;
use umya_spreadsheet::{CellRawValue, Workbook, Worksheet};

pub const MTC: &str = "Medium_Term_Cashflow";
pub const SCHEDULE: &str = "Schedule_of_Investments";
const KEPT: [&str; 2] = [MTC, SCHEDULE];

/// The defined names the extraction needs.
const REQUIRED: [&str; 7] = [
    "Cfg_BaseYear",
    "MTC_Title",
    "MTC_Years",
    "MTC_CashSection",
    "MTC_ComparisonSection",
    "Inv_Title",
    "Inv_Totals",
];

/// A name's target: its sheet and its rectangle as ((row, col), (row, col)).
#[derive(Clone, Debug)]
struct Target {
    sheet: String,
    start: (u32, u32),
    end: (u32, u32),
    address: String,
}

/// What [`extract`] did, for the command line to report.
#[derive(Debug, Default)]
pub struct Summary {
    pub output: PathBuf,
    pub formulas_kept: usize,
    pub formulas_to_values: usize,
}

/// The output path for a source file: `Spatac-20261121.xlsx` gives `InvestmentPlan-20261121.xlsx`
/// in the same directory. `None` when the source name carries no `YYYYMMDD` date.
pub fn default_output(source: &Path) -> Option<PathBuf> {
    static DATE: LazyLock<Regex> = LazyLock::new(|| Regex::new(r"(\d{8})").expect("valid pattern"));
    let name = source.file_name()?.to_string_lossy();
    let date = DATE.captures(&name)?.get(1)?.as_str().to_owned();
    Some(source.with_file_name(format!("InvestmentPlan-{date}.xlsx")))
}

/// Extracts the Investment Plan from `source` into `output`.
pub fn extract(source: &Path, output: &Path, generated_on: &str) -> Result<Summary, String> {
    let wb = umya_spreadsheet::reader::xlsx::read(source)
        .map_err(|e| format!("cannot read {}: {e}", source.display()))?;
    let names = defined_names(&wb);
    for n in REQUIRED {
        if !names.contains_key(n) {
            return Err(format!(
                "{} has no defined name {n}; is it a restructured Spatac workbook?",
                source.display()
            ));
        }
    }
    for sheet in KEPT {
        wb.sheet_by_name(sheet)
            .map_err(|_| format!("{} has no {sheet} tab", source.display()))?;
    }

    let mut summary = Summary {
        output: output.to_owned(),
        ..Summary::default()
    };
    let mut out = umya_spreadsheet::new_file();
    let placeholder = out.sheet_collection()[0].name().to_owned();

    // Cover.
    let base_year = cell_number(&wb, &names["Cfg_BaseYear"], 0, 0).unwrap_or(0.0);
    let source_name = source
        .file_name()
        .map_or_else(String::new, |f| f.to_string_lossy().into_owned());
    out.add_sheet(cover(&source_name, generated_on, base_year))
        .map_err(|e| e.to_string())?;

    // The two tabs, trimmed to what the plan needs.
    let mtc_keep = bounds(&[
        &names["MTC_Title"],
        &names["MTC_CashSection"],
        &names["MTC_ComparisonSection"],
    ]);
    let inv_keep = bounds(&[&names["Inv_Title"], &names["Inv_Totals"]]);
    let mut used_names = BTreeSet::new();
    for (sheet, keep) in [(MTC, mtc_keep), (SCHEDULE, inv_keep)] {
        let mut ws = wb.sheet_by_name(sheet).expect("checked above").clone();
        trim(&mut ws, keep);
        convert_formulas(&mut ws, &names, &mut used_names, &mut summary)?;
        lock(&mut ws);
        out.add_sheet(ws).map_err(|e| e.to_string())?;
    }

    out.remove_sheet_by_name(&placeholder)
        .map_err(|e| e.to_string())?;

    // Names the kept formulas still use, all of which point into the kept tabs.
    for n in used_names {
        let t = &names[&n];
        let mut dn = DefinedName::default();
        dn.set_name(n.clone());
        dn.set_address(t.address.clone());
        out.add_defined_names(dn);
    }
    for ws in out.sheet_collection_mut() {
        lock(ws);
    }
    out.set_active_sheet(0);

    umya_spreadsheet::writer::xlsx::write(&out, output)
        .map_err(|e| format!("cannot write {}: {e}", output.display()))?;
    Ok(summary)
}

fn defined_names(wb: &Workbook) -> BTreeMap<String, Target> {
    let mut out = BTreeMap::new();
    let all = wb
        .defined_names()
        .iter()
        .chain(wb.sheet_collection().iter().flat_map(|s| s.defined_names()));
    for dn in all {
        let address = dn.address();
        let tokens = formula::tokenize(&address);
        let [Token::Ref(r)] = tokens.as_slice() else {
            continue;
        };
        let Some(sheet) = r.sheet.clone() else {
            continue;
        };
        let Some((start, end)) = parse_range(&address) else {
            continue;
        };
        out.insert(
            dn.name().to_owned(),
            Target {
                sheet,
                start,
                end,
                address,
            },
        );
    }
    out
}

/// The rectangle covering every target, as ((row, col), (row, col)).
fn bounds(targets: &[&Target]) -> ((u32, u32), (u32, u32)) {
    let r1 = targets.iter().map(|t| t.start.0).min().unwrap_or(1);
    let c1 = targets.iter().map(|t| t.start.1).min().unwrap_or(1);
    let r2 = targets.iter().map(|t| t.end.0).max().unwrap_or(1);
    let c2 = targets.iter().map(|t| t.end.1).max().unwrap_or(1);
    ((r1.min(1), c1.min(1)), (r2, c2))
}

/// Removes every cell outside `keep`.
fn trim(ws: &mut Worksheet, keep: ((u32, u32), (u32, u32))) {
    let ((r1, c1), (r2, c2)) = keep;
    let outside: Vec<(u32, u32)> = ws
        .cells()
        .iter()
        .map(|c| (c.coordinate().col_num(), c.coordinate().row_num()))
        .filter(|&(c, r)| r < r1 || r > r2 || c < c1 || c > c2)
        .collect();
    for (c, r) in outside {
        ws.remove_cell((c, r));
    }
    ws.comments_mut().retain(|c| {
        let (r, col) = (c.coordinate().row_num(), c.coordinate().col_num());
        (r1..=r2).contains(&r) && (c1..=c2).contains(&col)
    });
}

static IDENT: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"[A-Za-z_][A-Za-z0-9_.]*").expect("valid pattern"));

/// Keeps formulas that stay within the kept tabs; replaces the others by their cached values.
fn convert_formulas(
    ws: &mut Worksheet,
    names: &BTreeMap<String, Target>,
    used: &mut BTreeSet<String>,
    summary: &mut Summary,
) -> Result<(), String> {
    let sheet = ws.name().to_owned();
    let coords: Vec<(u32, u32)> = ws
        .cells()
        .iter()
        .filter(|c| c.is_formula())
        .map(|c| (c.coordinate().col_num(), c.coordinate().row_num()))
        .collect();
    for coord in coords {
        let cell = ws.cell_mut(coord);
        let tokens = formula::tokenize(cell.formula());
        let mut external = false;
        let mut referenced = Vec::new();
        for token in &tokens {
            match token {
                Token::Ref(r) => {
                    if let Some(s) = &r.sheet
                        && !KEPT.contains(&s.as_str())
                    {
                        external = true;
                    }
                }
                Token::Text(text) => {
                    for m in IDENT.find_iter(&strip_strings(text)) {
                        if let Some(t) = names.get(m.as_str()) {
                            if KEPT.contains(&t.sheet.as_str()) {
                                referenced.push(m.as_str().to_owned());
                            } else {
                                external = true;
                            }
                        }
                    }
                }
            }
        }
        if !external {
            used.extend(referenced);
            summary.formulas_kept += 1;
            continue;
        }
        let raw = cell.cell_value().raw_value().clone();
        let value = cell.cell_value_mut();
        match raw {
            CellRawValue::Numeric(n) => {
                value.set_value_number(n);
            }
            CellRawValue::String(s) | CellRawValue::Lazy(s) => {
                value.set_value_string(s.to_string());
            }
            CellRawValue::RichText(r) => {
                value.set_value_string(r.text().into_owned());
            }
            CellRawValue::Bool(b) => {
                value.set_value_bool(b);
            }
            CellRawValue::Error(e) => {
                return Err(format!(
                    "{sheet}!{} holds the error {e}; recalculate and save the workbook first",
                    formula::cell_name(coord.1, coord.0)
                ));
            }
            CellRawValue::Empty => {
                value.set_value_string("");
            }
        }
        summary.formulas_to_values += 1;
    }
    Ok(())
}

/// Blanks out string literals, so that text inside quotes is not taken for a name.
fn strip_strings(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    let mut in_string = false;
    for c in text.chars() {
        if c == '"' {
            in_string = !in_string;
            out.push(' ');
        } else if in_string {
            out.push(' ');
        } else {
            out.push(c);
        }
    }
    out
}

/// Locks every cell and protects the sheet without a password.
fn lock(ws: &mut Worksheet) {
    for cell in ws.cells_mut() {
        cell.style_mut().protection_mut().set_locked(true);
    }
    ws.sheet_protection_mut().set_sheet(true);
}

fn cell_number(wb: &Workbook, t: &Target, dr: u32, dc: u32) -> Option<f64> {
    let ws = wb.sheet_by_name(&t.sheet).ok()?;
    ws.cell((t.start.1 + dc, t.start.0 + dr))?.value_number()
}

fn set_text(ws: &mut Worksheet, at: &str, text: &str, bold: bool) {
    let cell = ws.cell_mut(at);
    cell.set_value_string(text);
    if bold {
        cell.style_mut().font_mut().set_bold(true);
    }
}

fn cover(source: &str, generated_on: &str, base_year: f64) -> Worksheet {
    let mut ws = Worksheet::default();
    ws.set_name("Cover");
    set_text(&mut ws, "A1", "INVESTMENT PLAN — RESERVE FUND", true);
    ws.cell_mut("A1").style_mut().font_mut().set_size(14.0);
    set_text(
        &mut ws,
        "A2",
        "The Medium-Term Cash Flow schedule required by the Reserve Fund Investment Policy (s.3.2), \
         with the schedule of investments.",
        false,
    );
    let rows: [(&str, String); 3] = [
        ("Source workbook", source.to_owned()),
        ("Generated", generated_on.to_owned()),
        ("Base year", format!("{}", base_year as i64)),
    ];
    for (i, (label, value)) in rows.iter().enumerate() {
        let r = 4 + i as u32;
        set_text(&mut ws, &format!("A{r}"), label, true);
        set_text(&mut ws, &format!("B{r}"), value, false);
    }
    set_text(
        &mut ws,
        "A8",
        "Approved by the Board at the meeting of",
        true,
    );
    set_text(&mut ws, "B8", "____________________", false);
    set_text(&mut ws, "A9", "Minutes reference", true);
    set_text(&mut ws, "B9", "____________________", false);
    set_text(&mut ws, "A11", "Contents", true);
    let contents = [
        "Medium_Term_Cashflow — cash, expenditures, maturities and planned investments against the minimum target.",
        "Schedule_of_Investments — investments held and planned, with their cash flows by year.",
    ];
    for (i, line) in contents.iter().enumerate() {
        set_text(&mut ws, &format!("A{}", 12 + i), line, false);
    }
    set_text(&mut ws, "A15", "Changes from the Reserve Fund Study", true);
    set_text(
        &mut ws,
        "A16",
        "The expenditures in this plan differ from the Reserve Fund Study where the Board has changed \
         the timing or cost of a project. The reasons for those changes are described in the Spatac \
         workbooks for the current and previous years.",
        false,
    );
    ws.column_dimension_mut("A").set_width(42.0);
    ws.column_dimension_mut("B").set_width(60.0);
    ws
}

/// Today's date as `YYYY-MM-DD`, in the computer's own time zone: an evening run in Toronto is
/// dated that evening, not the next day as the UTC date would have it.
pub fn today() -> String {
    chrono::Local::now().format("%Y-%m-%d").to_string()
}

/// Checks a date given on the command line: `YYYY-MM-DD`, and a real day.
pub fn parse_date(text: &str) -> Result<String, String> {
    chrono::NaiveDate::parse_from_str(text, "%Y-%m-%d")
        .ok()
        .filter(|_| text.len() == 10)
        .map(|d| d.format("%Y-%m-%d").to_string())
        .ok_or_else(|| format!("{text} is not a date in the form YYYY-MM-DD"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn output_name_follows_the_source_date() {
        assert_eq!(
            default_output(Path::new("/x/Spatac-20261121.xlsx")),
            Some(PathBuf::from("/x/InvestmentPlan-20261121.xlsx"))
        );
        assert_eq!(default_output(Path::new("Spatac.xlsx")), None);
    }

    #[test]
    fn dates_are_checked() {
        assert_eq!(parse_date("2026-11-21").as_deref(), Ok("2026-11-21"));
        assert!(parse_date("2026-02-30").is_err());
        assert!(parse_date("2026-1-5").is_err());
        assert!(parse_date("21/11/2026").is_err());
    }

    #[test]
    fn today_is_a_date() {
        let t = today();
        assert_eq!(t.len(), 10);
        assert!(t.starts_with("20"));
    }
}
