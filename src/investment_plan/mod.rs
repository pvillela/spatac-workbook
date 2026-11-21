//! Investment Plan extraction.
//!
//! [`extract`] reads a Spatac workbook and writes the Investment Plan as a workbook of its own,
//! for the Board, the Treasurer and the broker (Reserve Fund Investment Policy s.3.2):
//!
//! - **Cover**: what the plan is, the source file, when it was generated, the base year, and
//!   lines to record the meeting that approved it.
//! - **Medium_Term_Cashflow**: the cash section and the RFS-versus-Revised comparison. The
//!   investment-balance comparison and notes are dropped.
//! - **Schedule_of_Investments**: held and planned investments with their cash flows by year.
//! - **Changes_from_RFS**: every item whose Revised amounts differ from the RFS within the
//!   horizon, with its reason; an item without a reason is flagged.
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
pub const EXPENDITURES: &str = "Expenditures";
const KEPT: [&str; 2] = [MTC, SCHEDULE];

/// The defined names the extraction needs, and why.
const REQUIRED: [&str; 14] = [
    "Cfg_BaseYear",
    "MTC_Title",
    "MTC_Years",
    "MTC_CashSection",
    "MTC_ComparisonSection",
    "Inv_Title",
    "Inv_Totals",
    "Exp_Item",
    "Exp_Element",
    "Exp_GLName",
    "Exp_RFS",
    "Exp_RFSYears",
    "Exp_Revised",
    "Exp_RevisedYears",
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
    pub changes: usize,
    pub changes_without_reason: usize,
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
    for sheet in KEPT.iter().chain([&EXPENDITURES]) {
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

    // Changes from the RFS.
    let changes = changes_from_rfs(&wb, &names)?;
    summary.changes = changes.len();
    summary.changes_without_reason = changes.iter().filter(|c| c.reason.is_empty()).count();
    out.add_sheet(changes_sheet(&changes, base_year))
        .map_err(|e| e.to_string())?;

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

fn cell_text(wb: &Workbook, t: &Target, dr: u32, dc: u32) -> String {
    wb.sheet_by_name(&t.sheet)
        .ok()
        .and_then(|ws| ws.cell((t.start.1 + dc, t.start.0 + dr)))
        .map(|c| c.value().trim().to_owned())
        .unwrap_or_default()
}

/// One item whose Revised amounts differ from the RFS within the horizon.
#[derive(Debug)]
pub struct Change {
    pub item: String,
    pub element: String,
    /// (year, amount) where the RFS amount is not the Revised one and not zero.
    pub rfs: Vec<(i32, f64)>,
    pub revised: Vec<(i32, f64)>,
    /// Revised minus RFS over the horizon.
    pub net: f64,
    pub reason: String,
}

fn changes_from_rfs(
    wb: &Workbook,
    names: &BTreeMap<String, Target>,
) -> Result<Vec<Change>, String> {
    let horizon: Vec<i32> = {
        let t = &names["MTC_Years"];
        (0..=(t.end.1 - t.start.1))
            .filter_map(|k| cell_number(wb, t, 0, k).map(|y| y as i32))
            .collect()
    };
    let year_cols = |years_name: &str| -> BTreeMap<i32, u32> {
        let t = &names[years_name];
        (0..=(t.end.1 - t.start.1))
            .filter_map(|k| cell_number(wb, t, 0, k).map(|y| (y as i32, k)))
            .collect()
    };
    let (rfs_cols, rev_cols) = (year_cols("Exp_RFSYears"), year_cols("Exp_RevisedYears"));
    let (item, element, gl_name) = (
        &names["Exp_Item"],
        &names["Exp_Element"],
        &names["Exp_GLName"],
    );
    let (rfs, revised) = (&names["Exp_RFS"], &names["Exp_Revised"]);
    let reason = names.get("Exp_Reason");
    let rows = item.end.0 - item.start.0 + 1;
    let mut out = Vec::new();
    for i in 0..rows {
        let mut a = Vec::new();
        let mut b = Vec::new();
        let mut net = 0.0;
        for &y in &horizon {
            let x = rfs_cols
                .get(&y)
                .and_then(|&k| cell_number(wb, rfs, i, k))
                .unwrap_or(0.0);
            let z = rev_cols
                .get(&y)
                .and_then(|&k| cell_number(wb, revised, i, k))
                .unwrap_or(0.0);
            if (x - z).abs() >= 0.5 {
                if x.abs() >= 0.5 {
                    a.push((y, x));
                }
                if z.abs() >= 0.5 {
                    b.push((y, z));
                }
                net += z - x;
            }
        }
        if a.is_empty() && b.is_empty() {
            continue;
        }
        let mut el = cell_text(wb, element, i, 0);
        if el.is_empty() {
            el = cell_text(wb, gl_name, i, 0);
        }
        out.push(Change {
            item: cell_text(wb, item, i, 0),
            element: el,
            rfs: a,
            revised: b,
            net,
            reason: reason.map(|t| cell_text(wb, t, i, 0)).unwrap_or_default(),
        });
    }
    Ok(out)
}

fn amounts(list: &[(i32, f64)]) -> String {
    if list.is_empty() {
        return "—".to_owned();
    }
    list.iter()
        .map(|(y, v)| format!("{y}: {}", thousands(*v)))
        .collect::<Vec<_>>()
        .join("; ")
}

/// Whole dollars with thousands separators.
fn thousands(v: f64) -> String {
    let n = v.round() as i64;
    let digits = n.unsigned_abs().to_string();
    let mut s = String::new();
    for (i, ch) in digits.chars().enumerate() {
        if i > 0 && (digits.len() - i).is_multiple_of(3) {
            s.push(',');
        }
        s.push(ch);
    }
    if n < 0 { format!("-{s}") } else { s }
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
         with the schedule of investments and the Board's changes to the Reserve Fund Study.",
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
        "Changes_from_RFS — items whose amounts or timing differ from the Reserve Fund Study, with reasons.",
    ];
    for (i, line) in contents.iter().enumerate() {
        set_text(&mut ws, &format!("A{}", 12 + i), line, false);
    }
    ws.column_dimension_mut("A").set_width(42.0);
    ws.column_dimension_mut("B").set_width(60.0);
    ws
}

fn changes_sheet(changes: &[Change], base_year: f64) -> Worksheet {
    let mut ws = Worksheet::default();
    ws.set_name("Changes_from_RFS");
    set_text(&mut ws, "A1", "CHANGES FROM THE RESERVE FUND STUDY", true);
    ws.cell_mut("A1").style_mut().font_mut().set_size(14.0);
    set_text(
        &mut ws,
        "A2",
        &format!(
            "Items whose Revised amounts differ from the RFS in {}–{}: the RFS amounts and years \
             replaced, the Revised ones, and the reason.",
            base_year as i64,
            base_year as i64 + 5
        ),
        false,
    );
    let headers = [
        "Item",
        "Building element",
        "RFS",
        "Revised",
        "Net change",
        "Reason",
    ];
    for (k, h) in headers.iter().enumerate() {
        let at = format!("{}4", formula::column_letters(k as u32 + 1));
        set_text(&mut ws, &at, h, true);
    }
    for (i, c) in changes.iter().enumerate() {
        let r = 5 + i as u32;
        set_text(&mut ws, &format!("A{r}"), &c.item, false);
        set_text(&mut ws, &format!("B{r}"), &c.element, false);
        set_text(&mut ws, &format!("C{r}"), &amounts(&c.rfs), false);
        set_text(&mut ws, &format!("D{r}"), &amounts(&c.revised), false);
        let net = ws.cell_mut(format!("E{r}").as_str());
        net.set_value_number(c.net.round());
        net.style_mut()
            .number_format_mut()
            .set_format_code("#,##0;-#,##0");
        let reason = if c.reason.is_empty() {
            "REASON MISSING"
        } else {
            &c.reason
        };
        set_text(&mut ws, &format!("F{r}"), reason, c.reason.is_empty());
    }
    let total = 5 + changes.len() as u32 + 1;
    set_text(
        &mut ws,
        &format!("A{total}"),
        "Net change over the horizon",
        true,
    );
    let net: f64 = changes.iter().map(|c| c.net).sum();
    let cell = ws.cell_mut(format!("E{total}").as_str());
    cell.set_value_number(net.round());
    cell.style_mut()
        .number_format_mut()
        .set_format_code("#,##0;-#,##0");
    cell.style_mut().font_mut().set_bold(true);
    for (c, w) in [
        ("A", 10.0),
        ("B", 40.0),
        ("C", 34.0),
        ("D", 34.0),
        ("E", 14.0),
        ("F", 60.0),
    ] {
        ws.column_dimension_mut(c).set_width(w);
    }
    ws
}

/// Today's date as `YYYY-MM-DD`, from the system clock (UTC).
pub fn today() -> String {
    let secs = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map_or(0, |d| d.as_secs());
    let days = (secs / 86_400) as i64;
    // Howard Hinnant's civil-from-days.
    let z = days + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = if m <= 2 { y + 1 } else { y };
    format!("{y:04}-{m:02}-{d:02}")
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
    fn thousands_separators() {
        assert_eq!(thousands(1234567.4), "1,234,567");
        assert_eq!(thousands(-587.0), "-587");
        assert_eq!(thousands(999.0), "999");
    }

    #[test]
    fn today_is_a_date() {
        let t = today();
        assert_eq!(t.len(), 10);
        assert!(t.starts_with("20"));
    }
}
