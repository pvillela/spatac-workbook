//! The parts of a workbook the diff compares, read out of umya-spreadsheet into plain data.

use super::formula::{self, Token};
use std::collections::hash_map::DefaultHasher;
use std::collections::{BTreeMap, BTreeSet};
use std::hash::{Hash, Hasher};
use std::path::Path;
use umya_spreadsheet::{CellRawValue, Worksheet};

/// A cell's constant value, or a formula's cached result.
#[derive(Clone, Debug)]
pub enum Value {
    Number(f64),
    Text(String),
    Bool(bool),
    Error(String),
}

impl Value {
    /// Numbers compare equal within a relative 1e-9, which absorbs the different number of
    /// digits Excel and LibreOffice write for the same value, or an absolute 1e-6, which absorbs
    /// the rounding residue of a sum that should be zero. Amounts here are dollars, so neither
    /// hides a real change.
    pub fn same(&self, other: &Value) -> bool {
        match (self, other) {
            (Value::Number(a), Value::Number(b)) => {
                a == b || (a - b).abs() <= (1e-9 * a.abs().max(b.abs())).max(1e-6)
            }
            (Value::Text(a), Value::Text(b)) => a == b,
            (Value::Bool(a), Value::Bool(b)) => a == b,
            (Value::Error(a), Value::Error(b)) => a == b,
            _ => false,
        }
    }

    pub fn display(&self) -> String {
        match self {
            Value::Number(n) => format!("{n}"),
            Value::Text(t) => format!("\"{t}\""),
            Value::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_owned(),
            Value::Error(e) => e.clone(),
        }
    }

    fn signature(&self) -> String {
        match self {
            // Rounded, so that values that compare equal also hash equal.
            Value::Number(n) => format!("n{:.9e}", n),
            Value::Text(t) => format!("s{t}"),
            Value::Bool(b) => format!("b{b}"),
            Value::Error(e) => format!("e{e}"),
        }
    }
}

#[derive(Clone, Debug)]
pub enum Content {
    Empty,
    Value(Value),
    Formula {
        /// Formula text without the leading `=`.
        text: String,
        tokens: Vec<Token>,
        /// The range of a legacy array formula, entered once and shown across the range.
        array: Option<String>,
    },
    /// A cell inside an array formula's range, other than the one holding the formula.
    ArrayMember,
}

#[derive(Clone, Debug)]
pub struct Cell {
    pub content: Content,
    /// A formula's cached result.
    pub cached: Option<Value>,
    pub comment: Option<String>,
    pub unlocked: bool,
    /// Present only when formats are being compared.
    pub style: Option<u64>,
}

impl Cell {
    /// A signature of what the cell says, independent of where it is: formulas in relative form.
    pub fn content_signature(&self, row: u32, col: u32) -> String {
        match &self.content {
            Content::Empty => String::new(),
            Content::Value(v) => v.signature(),
            Content::Formula { tokens, .. } => {
                format!("={}", formula::relative_form(tokens, row, col))
            }
            Content::ArrayMember => "{array}".to_owned(),
        }
    }

    pub fn has_content(&self) -> bool {
        !matches!(self.content, Content::Empty)
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct Protection {
    pub has_password: bool,
}

#[derive(Clone, Debug)]
pub struct Sheet {
    pub name: String,
    /// `visible`, `hidden` or `veryHidden`.
    pub state: String,
    /// Keyed by (row, column), both 1-based.
    pub cells: BTreeMap<(u32, u32), Cell>,
    /// The last row and column holding content, a comment or an unlocked cell.
    pub max_row: u32,
    pub max_col: u32,
    pub merges: Vec<String>,
    pub conditional_formats: Vec<ConditionalFormat>,
    pub validations: Vec<Validation>,
    pub protection: Option<Protection>,
    pub hidden_rows: BTreeSet<u32>,
    pub hidden_cols: BTreeSet<u32>,
    /// Frozen (rows, columns).
    pub frozen: Option<(u32, u32)>,
    pub row_heights: BTreeMap<u32, f64>,
    pub col_widths: BTreeMap<u32, f64>,
}

#[derive(Clone, Debug)]
pub struct ConditionalFormat {
    pub ranges: Vec<String>,
    /// The rule without its ranges, formula tokens kept separately so they can be remapped.
    pub rule: String,
    pub formula: Option<Vec<Token>>,
}

#[derive(Clone, Debug)]
pub struct Validation {
    pub ranges: Vec<String>,
    pub rule: String,
    pub formula1: Vec<Token>,
    pub formula2: Vec<Token>,
}

#[derive(Clone, Debug)]
pub struct Name {
    pub name: String,
    /// The sheet a local name belongs to.
    pub scope: Option<String>,
    pub refers_to: Vec<Token>,
}

#[derive(Clone, Debug)]
pub struct Book {
    pub sheets: Vec<Sheet>,
    pub names: Vec<Name>,
}

/// Reads a workbook. `with_styles` also records a signature of every cell's style, which costs
/// time and is needed only for `--formats`.
pub fn load(path: &Path, with_styles: bool) -> Result<Book, String> {
    let wb = umya_spreadsheet::reader::xlsx::read(path)
        .map_err(|e| format!("cannot read {}: {e}", path.display()))?;
    let sheet_names: Vec<String> = wb
        .sheet_collection()
        .iter()
        .map(|s| s.name().to_owned())
        .collect();
    let mut names = Vec::new();
    for dn in wb.defined_names() {
        names.push(Name {
            name: dn.name().to_owned(),
            scope: dn
                .has_local_sheet_id()
                .then(|| sheet_names.get(dn.local_sheet_id() as usize).cloned())
                .flatten(),
            refers_to: formula::tokenize(&dn.address()),
        });
    }
    let mut sheets = Vec::new();
    for ws in wb.sheet_collection() {
        // umya files a workbook-level name under the sheet it refers to; only a name carrying a
        // local sheet id is really scoped to that sheet.
        for dn in ws.defined_names() {
            names.push(Name {
                name: dn.name().to_owned(),
                scope: dn.has_local_sheet_id().then(|| ws.name().to_owned()),
                refers_to: formula::tokenize(&dn.address()),
            });
        }
        sheets.push(load_sheet(ws, with_styles));
    }
    Ok(Book { sheets, names })
}

fn load_sheet(ws: &Worksheet, with_styles: bool) -> Sheet {
    let mut cells: BTreeMap<(u32, u32), Cell> = BTreeMap::new();
    let mut arrays: Vec<(u32, u32, String)> = Vec::new();
    for c in ws.cells() {
        let row = c.coordinate().row_num();
        let col = c.coordinate().col_num();
        let unlocked = c.style().protection().is_some_and(|p| !p.locked());
        let (content, cached) = if c.is_formula() {
            let fo = c.formula_obj();
            let array = fo
                .filter(|f| {
                    matches!(
                        f.formula_type(),
                        umya_spreadsheet::structs::CellFormulaValues::Array
                    )
                })
                .map(|f| f.reference().to_owned())
                .filter(|r| !r.is_empty());
            if let Some(r) = &array {
                arrays.push((row, col, r.clone()));
            }
            let text = c.formula().to_owned();
            (
                Content::Formula {
                    tokens: formula::tokenize(&text),
                    text,
                    array,
                },
                raw_value(c.cell_value().raw_value()),
            )
        } else {
            match raw_value(c.cell_value().raw_value()) {
                Some(v) => (Content::Value(v), None),
                None => (Content::Empty, None),
            }
        };
        let style = with_styles.then(|| {
            let mut h = DefaultHasher::new();
            format!("{:?}", c.style()).hash(&mut h);
            h.finish()
        });
        if !with_styles && matches!(content, Content::Empty) && !unlocked {
            continue;
        }
        cells.insert(
            (row, col),
            Cell {
                content,
                cached,
                comment: None,
                unlocked,
                style,
            },
        );
    }
    // The cells an array formula spans, other than its own, belong to it.
    for (row, col, range) in arrays {
        if let Some(((r1, c1), (r2, c2))) = parse_range(&range) {
            for r in r1..=r2 {
                for c in c1..=c2 {
                    if (r, c) == (row, col) {
                        continue;
                    }
                    let cell = cells.entry((r, c)).or_insert_with(|| Cell {
                        content: Content::Empty,
                        cached: None,
                        comment: None,
                        unlocked: false,
                        style: None,
                    });
                    if !matches!(cell.content, Content::Formula { .. }) {
                        cell.cached =
                            match std::mem::replace(&mut cell.content, Content::ArrayMember) {
                                Content::Value(v) => Some(v),
                                _ => None,
                            };
                    }
                }
            }
        }
    }
    for comment in ws.comments() {
        let row = comment.coordinate().row_num();
        let col = comment.coordinate().col_num();
        let text = comment
            .text()
            .rich_text()
            .map(|r| r.text().into_owned())
            .or_else(|| comment.text().text().map(|t| t.value().to_owned()))
            .unwrap_or_default();
        cells
            .entry((row, col))
            .or_insert_with(|| Cell {
                content: Content::Empty,
                cached: None,
                comment: None,
                unlocked: false,
                style: None,
            })
            .comment = Some(text);
    }
    let (mut max_row, mut max_col) = (0, 0);
    for (&(r, c), cell) in &cells {
        if cell.has_content() || cell.comment.is_some() || cell.unlocked {
            max_row = max_row.max(r);
            max_col = max_col.max(c);
        }
    }

    let merges = ws.merge_cells().iter().map(|r| r.range()).collect();

    let mut conditional_formats = Vec::new();
    for cf in ws.conditional_formatting_collection() {
        let ranges: Vec<String> = cf
            .sequence_of_references()
            .range_collection()
            .iter()
            .map(|r| r.range())
            .collect();
        for rule in cf.conditional_collection() {
            let style = rule.style().map(|s| {
                let mut h = DefaultHasher::new();
                format!("{s:?}").hash(&mut h);
                format!(" style#{:08x}", h.finish() as u32)
            });
            let mut desc = format!("{:?}", rule.get_type());
            // The operator means something only for cell-value rules; elsewhere it is a default.
            if desc == "CellIs" {
                desc.push_str(&format!(" {:?}", rule.operator()));
            }
            if !rule.text().is_empty() {
                desc.push_str(&format!(" text=\"{}\"", rule.text()));
            }
            desc.push_str(&style.unwrap_or_default());
            conditional_formats.push(ConditionalFormat {
                ranges: ranges.clone(),
                rule: desc,
                formula: rule.formula().map(|f| formula::tokenize(&f.address_str())),
            });
        }
    }

    let mut validations = Vec::new();
    if let Some(dvs) = ws.data_validations() {
        for dv in dvs.data_validation_list() {
            validations.push(Validation {
                ranges: dv
                    .sequence_of_references()
                    .range_collection()
                    .iter()
                    .map(|r| r.range())
                    .collect(),
                rule: format!(
                    "{:?} {:?}{}",
                    dv.get_type(),
                    dv.operator(),
                    if dv.allow_blank() { " allow-blank" } else { "" }
                ),
                formula1: formula::tokenize(dv.formula1()),
                formula2: formula::tokenize(dv.formula2()),
            });
        }
    }

    let protection = ws
        .sheet_protection()
        .filter(|p| p.sheet())
        .map(|p| Protection {
            has_password: !p.hash_value().is_empty() || !p.password_raw().is_empty(),
        });

    let mut hidden_rows = BTreeSet::new();
    let mut row_heights = BTreeMap::new();
    for row in ws.row_dimensions() {
        if row.hidden() {
            hidden_rows.insert(row.row_num());
        }
        if row.custom_height() {
            row_heights.insert(row.row_num(), row.height());
        }
    }
    let mut hidden_cols = BTreeSet::new();
    let mut col_widths = BTreeMap::new();
    for col in ws.column_dimensions() {
        if col.hidden() {
            hidden_cols.insert(col.col_num());
        }
        col_widths.insert(col.col_num(), col.width());
    }

    let frozen = ws
        .sheets_views()
        .sheet_view_list()
        .first()
        .and_then(|v| v.pane())
        .filter(|p| format!("{:?}", p.state()).starts_with("Frozen"))
        .map(|p| (p.vertical_split() as u32, p.horizontal_split() as u32));

    let state = match ws.sheet_state() {
        "" => "visible".to_owned(),
        s => s.to_owned(),
    };

    Sheet {
        name: ws.name().to_owned(),
        state,
        cells,
        max_row,
        max_col,
        merges,
        conditional_formats,
        validations,
        protection,
        hidden_rows,
        hidden_cols,
        frozen,
        row_heights,
        col_widths,
    }
}

fn raw_value(raw: &CellRawValue) -> Option<Value> {
    match raw {
        CellRawValue::Numeric(n) => Some(Value::Number(*n)),
        CellRawValue::String(s) | CellRawValue::Lazy(s) => Some(Value::Text(s.to_string())),
        CellRawValue::RichText(r) => Some(Value::Text(r.text().into_owned())),
        CellRawValue::Bool(b) => Some(Value::Bool(*b)),
        CellRawValue::Error(e) => Some(Value::Error(e.to_string())),
        CellRawValue::Empty => None,
    }
}

/// `B2:D5` as ((2, 2), (5, 4)): (row, column) pairs.
pub fn parse_range(range: &str) -> Option<((u32, u32), (u32, u32))> {
    let tokens = formula::tokenize(range);
    let [Token::Ref(r)] = tokens.as_slice() else {
        return None;
    };
    let start = (r.start.row?, r.start.col?);
    let end = match &r.end {
        Some(e) => (e.row?, e.col?),
        None => start,
    };
    Some((start, end))
}
