//! Integrity checks of a Spatac workbook (design 14.10).
//!
//! [`check`] reads a saved workbook, without recalculating it, and lists every problem it finds:
//!
//! - **Missing formulas**: a formula column with a gap. Each column is found by its defined name,
//!   so the check follows inserted rows. On Adjustments, a row counts when its Item cell is an
//!   input (unlocked), so a row inserted into a block and given the input look is checked too.
//! - **Names**: a name the workbook relies on that is missing, or one that points at `#REF!`.
//! - **Errors**: a cell whose stored result is an error (`#N/A`, `#REF!`…), as Excel or Calc
//!   last calculated it.
//! - **Validations**: an Adjustments input row without its dropdowns or its amount rule.
//! - **Protection**: an unprotected tab (the Spatac tabs and the three RFS tabs), a formula cell
//!   that is unlocked, or an Adjustments input cell that is locked.

use crate::diff::formula::{self, Token};
use crate::diff::model::parse_range;
use crate::diff::protection::unlocked_cells;
use std::collections::{BTreeMap, BTreeSet};
use std::fmt;
use std::path::Path;
use umya_spreadsheet::structs::DataValidationValues;
use umya_spreadsheet::{CellRawValue, Workbook, Worksheet};

/// The tabs Spatac maintains, all protected.
const SPATAC_TABS: [&str; 7] = [
    "README",
    "Expenditures",
    "Adjustments",
    "Medium_Term_Cashflow",
    "Schedule_of_Investments",
    "Projected_Balances",
    "Settings",
];

/// The vendor's tabs, protected too (design 17.6): they are unprotected only to paste a new RFS.
const VENDOR_TABS: [&str; 3] = ["RFS_info", "RFS_fvalues", "RFS_cashflow"];

/// Which cells of a named range must hold formulas.
#[derive(Clone, Copy)]
enum Cells {
    /// Every cell.
    All,
    /// All but the first row (the header) and the last (▼ End).
    Items,
    /// All but the last column (▶ End).
    Years,
    /// The rows whose cell in the named column is an input.
    InputRows(&'static str),
}

/// The formula columns, by name.
const FORMULAS: &[(&str, Cells)] = &[
    ("Exp_Key", Cells::Items),
    ("Exp_Key2", Cells::Items),
    ("Exp_Check", Cells::Items),
    ("Exp_AdjLabel", Cells::Items),
    ("Exp_Current", Cells::Items),
    ("Exp_NextYear", Cells::Items),
    ("Exp_Adjusted", Cells::Items),
    ("Adj_FromAmount", Cells::InputRows("Adj_Item")),
    ("Adj_ToAmount", Cells::InputRows("Adj_Item")),
    ("Adj_Check", Cells::InputRows("Adj_Item")),
    ("MTC_Contributions", Cells::Years),
    ("MTC_Expenditures", Cells::Years),
    ("MTC_Maturities", Cells::Years),
    ("MTC_PlannedMaturities", Cells::Years),
    ("MTC_NewInvTotal", Cells::Years),
    ("MTC_ClosingCash", Cells::Years),
    ("MTC_MinCash", Cells::Years),
    ("MTC_RFSExpenditures", Cells::Years),
    ("MTC_RevisedExpenditures", Cells::Years),
    ("MTC_ExpVariance", Cells::Years),
    ("PB_Year", Cells::All),
    ("PB_RFSOpening", Cells::All),
    ("PB_RFSCost", Cells::All),
    ("PB_RFSContrib", Cells::All),
    ("PB_RFSInterest", Cells::All),
    ("PB_RFSBalance", Cells::All),
    ("PB_RFSDifference", Cells::All),
    ("PB_Opening", Cells::All),
    ("PB_Cost", Cells::All),
    ("PB_Contributions", Cells::All),
    ("PB_Interest", Cells::All),
    ("PB_Balance", Cells::All),
    ("PB_Variance", Cells::All),
    ("PB_Inflator", Cells::All),
];

/// Other names the workbook relies on.
const NAMES: &[&str] = &[
    "Cfg_BaseYear",
    "Cfg_ShortInflationYears",
    "Cfg_ShortInterestYears",
    "Cfg_YearsBetweenRFS",
    "RFS_Fv_Items",
    "Exp_AdjustedYears",
    "Adj_Item",
    "Adj_From",
    "Adj_To",
    "Adj_Amount",
    "Adj_Reason",
    "Adj_YearChoices",
    "MTC_Years",
];

/// The Adjustments input columns and the validation each needs.
const ADJ_VALIDATIONS: [(u32, DataValidationValues); 4] = [
    (1, DataValidationValues::List),
    (2, DataValidationValues::List),
    (3, DataValidationValues::List),
    (4, DataValidationValues::Custom),
];

/// One problem, at a cell, a tab or a name. Problems sort by tab, then row, then column.
#[derive(Debug, PartialEq, Eq, PartialOrd, Ord)]
pub struct Problem {
    sheet: String,
    row: u32,
    col: u32,
    pub place: String,
    pub message: String,
}

impl Problem {
    fn at(sheet: &str, row: u32, col: u32, message: &str) -> Self {
        Problem {
            sheet: sheet.to_owned(),
            row,
            col,
            place: format!("{sheet}!{}", formula::cell_name(row, col)),
            message: message.to_owned(),
        }
    }

    /// A problem with a tab or a name as a whole.
    fn whole(place: &str, message: &str) -> Self {
        Problem {
            sheet: place.to_owned(),
            row: 0,
            col: 0,
            place: place.to_owned(),
            message: message.to_owned(),
        }
    }
}

impl fmt::Display for Problem {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}: {}", self.place, self.message)
    }
}

type Rect = ((u32, u32), (u32, u32));

struct Context<'a> {
    wb: &'a Workbook,
    names: BTreeMap<String, (String, Rect)>,
    unlocked: crate::diff::protection::Unlocked,
    problems: BTreeSet<Problem>,
}

impl Context<'_> {
    fn add(&mut self, sheet: &str, row: u32, col: u32, message: &str) {
        self.problems.insert(Problem::at(sheet, row, col, message));
    }

    fn add_name(&mut self, name: &str, message: &str) {
        self.problems.insert(Problem::whole(name, message));
    }

    fn is_unlocked(&self, sheet: &str, row: u32, col: u32) -> bool {
        self.unlocked
            .get(sheet)
            .is_some_and(|s| s.contains(&(row, col)))
    }
}

/// Checks the workbook at `path` and returns every problem found, sorted by place.
pub fn check(path: &Path) -> Result<Vec<Problem>, String> {
    let wb = umya_spreadsheet::reader::xlsx::read(path)
        .map_err(|e| format!("cannot read {}: {e}", path.display()))?;
    let mut cx = Context {
        names: BTreeMap::new(),
        unlocked: unlocked_cells(path)?,
        problems: BTreeSet::new(),
        wb: &wb,
    };
    names(&mut cx);
    for &(name, cells) in FORMULAS {
        formulas(&mut cx, name, cells);
    }
    validations(&mut cx);
    for ws in wb.sheet_collection() {
        errors(&mut cx, ws);
    }
    for tab in SPATAC_TABS.into_iter().chain(VENDOR_TABS) {
        match wb.sheet_by_name(tab) {
            Ok(ws) if ws.sheet_protection().is_some_and(|p| p.sheet()) => {}
            Ok(_) => {
                cx.problems
                    .insert(Problem::whole(tab, "the tab is not protected"));
            }
            Err(_) => {
                cx.problems
                    .insert(Problem::whole(tab, "the tab is missing"));
            }
        }
    }
    Ok(cx.problems.into_iter().collect())
}

fn names(cx: &mut Context) {
    let all = cx.wb.defined_names().iter().chain(
        cx.wb
            .sheet_collection()
            .iter()
            .flat_map(|s| s.defined_names()),
    );
    let mut broken = Vec::new();
    for dn in all {
        let address = dn.address();
        if address.contains("#REF!") {
            broken.push(dn.name().to_owned());
            continue;
        }
        let tokens = formula::tokenize(&address);
        if let [Token::Ref(r)] = tokens.as_slice()
            && let Some(sheet) = r.sheet.clone()
            && let Some(rect) = parse_range(&address)
        {
            cx.names.insert(dn.name().to_owned(), (sheet, rect));
        }
    }
    for n in broken {
        cx.add_name(&n, "the name points at #REF!");
    }
    let required = FORMULAS
        .iter()
        .map(|(n, _)| *n)
        .chain(NAMES.iter().copied());
    let missing: Vec<&str> = required.filter(|n| !cx.names.contains_key(*n)).collect();
    for n in missing {
        cx.add_name(n, "the name is missing");
    }
}

fn formulas(cx: &mut Context, name: &str, cells: Cells) {
    let Some((sheet, ((r1, c1), (r2, c2)))) = cx.names.get(name).cloned() else {
        return;
    };
    let Ok(ws) = cx.wb.sheet_by_name(&sheet) else {
        return;
    };
    let (rows, cols) = match cells {
        Cells::All | Cells::InputRows(_) => (r1..=r2, c1..=c2),
        Cells::Items => (r1 + 1..=r2.saturating_sub(1), c1..=c2),
        Cells::Years => (r1..=r2, c1..=c2.saturating_sub(1)),
    };
    let input_col = match cells {
        Cells::InputRows(n) => cx.names.get(n).map(|(_, ((_, c), _))| *c),
        _ => None,
    };
    let mut found = Vec::new();
    for r in rows {
        if let Some(ic) = input_col
            && !cx.is_unlocked(&sheet, r, ic)
        {
            continue;
        }
        for c in cols.clone() {
            let is_formula = ws.cell((c, r)).is_some_and(|cell| cell.is_formula());
            if !is_formula {
                found.push((r, c, "a formula is missing (copy it from the row above)"));
            } else if cx.is_unlocked(&sheet, r, c) {
                found.push((r, c, "a formula cell is unlocked"));
            }
        }
    }
    for (r, c, m) in found {
        cx.add(&sheet, r, c, m);
    }
}

fn validations(cx: &mut Context) {
    let Some((sheet, ((r1, item_col), (r2, _)))) = cx.names.get("Adj_Item").cloned() else {
        return;
    };
    let Ok(ws) = cx.wb.sheet_by_name(&sheet) else {
        return;
    };
    let mut covered: BTreeMap<(u32, u32), Vec<DataValidationValues>> = BTreeMap::new();
    if let Some(dvs) = ws.data_validations() {
        for dv in dvs.data_validation_list() {
            for range in dv.sequence_of_references().range_collection() {
                let Some(((a, b), (c, d))) = parse_range(&range.range()) else {
                    continue;
                };
                for r in a..=c {
                    for col in b..=d {
                        covered
                            .entry((r, col))
                            .or_default()
                            .push(dv.get_type().clone());
                    }
                }
            }
        }
    }
    let mut found = Vec::new();
    for r in r1..=r2 {
        if !cx.is_unlocked(&sheet, r, item_col) {
            continue;
        }
        for (offset, kind) in &ADJ_VALIDATIONS {
            let col = item_col + offset - 1;
            if !covered.get(&(r, col)).is_some_and(|k| k.contains(kind)) {
                found.push((
                    r,
                    col,
                    "the input has lost its validation (copy the row above)",
                ));
            }
            if !cx.is_unlocked(&sheet, r, col) {
                found.push((r, col, "an input cell is locked"));
            }
        }
        let reason = item_col + 4;
        if !cx.is_unlocked(&sheet, r, reason) {
            found.push((r, reason, "an input cell is locked"));
        }
    }
    for (r, c, m) in found {
        cx.add(&sheet, r, c, m);
    }
}

fn errors(cx: &mut Context, ws: &Worksheet) {
    let found: Vec<(u32, u32, String)> = ws
        .cells()
        .iter()
        .filter_map(|c| match c.cell_value().raw_value() {
            CellRawValue::Error(e) => Some((
                c.coordinate().row_num(),
                c.coordinate().col_num(),
                format!("the cell shows the error {e}"),
            )),
            _ => None,
        })
        .collect();
    for (r, c, m) in found {
        cx.add(ws.name(), r, c, &m);
    }
}
