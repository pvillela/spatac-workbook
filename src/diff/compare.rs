//! Compares two loaded workbooks and builds the [`Report`].

use super::align::{Alignment, align, align_labelled};
use super::formula::{self, Remap, SheetMap, cell_name, column_letters};
use super::model::{Book, Cell, Content, Sheet, Value};
use similar::{Algorithm, DiffOp, capture_diff_slices};
use std::collections::hash_map::DefaultHasher;
use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::hash::{Hash, Hasher};

/// What to compare beyond the defaults.
#[derive(Clone, Copy, Debug, Default)]
pub struct Options {
    /// Also report changed cached results of formulas that did not change.
    pub values: bool,
    /// Also report changed cell styles, column widths and row heights.
    pub formats: bool,
}

#[derive(Debug, Default)]
pub struct Report {
    pub added_sheets: Vec<String>,
    pub deleted_sheets: Vec<String>,
    /// (old name, new name).
    pub renamed_sheets: Vec<(String, String)>,
    /// Tabs whose position changed relative to the others: (name, old position, new position),
    /// positions 1-based among the tabs present in both versions.
    pub reordered_sheets: Vec<(String, usize, usize)>,
    /// The full tab order, old then new, when it changed.
    pub order: Option<(Vec<String>, Vec<String>)>,
    pub names: Vec<NameChange>,
    /// Tabs present in both versions that changed, in new order.
    pub sheets: Vec<SheetReport>,
    /// Tabs present in both versions that did not change.
    pub unchanged_sheets: Vec<String>,
}

impl Report {
    pub fn is_empty(&self) -> bool {
        self.added_sheets.is_empty()
            && self.deleted_sheets.is_empty()
            && self.renamed_sheets.is_empty()
            && self.order.is_none()
            && self.names.is_empty()
            && self.sheets.is_empty()
    }
}

#[derive(Debug)]
pub enum NameChange {
    Added {
        name: String,
        refers_to: String,
    },
    Removed {
        name: String,
        refers_to: String,
    },
    Changed {
        name: String,
        before: String,
        after: String,
    },
}

#[derive(Debug, Default)]
pub struct SheetReport {
    pub name: String,
    pub renamed_from: Option<String>,
    /// Structural changes, one sentence each, in Markdown.
    pub structure: Vec<String>,
    pub cells: Vec<CellChange>,
    /// Changed formula results (`--values`).
    pub values: Vec<CellChange>,
    /// Changed formats (`--formats`), one sentence each.
    pub formats: Vec<String>,
}

impl SheetReport {
    fn is_empty(&self) -> bool {
        self.structure.is_empty()
            && self.cells.is_empty()
            && self.values.is_empty()
            && self.formats.is_empty()
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ChangeKind {
    Added,
    Removed,
    Changed,
    CommentAdded,
    CommentRemoved,
    CommentChanged,
    Value,
}

impl ChangeKind {
    pub fn label(self) -> &'static str {
        match self {
            ChangeKind::Added => "added",
            ChangeKind::Removed => "removed",
            ChangeKind::Changed => "changed",
            ChangeKind::CommentAdded => "comment added",
            ChangeKind::CommentRemoved => "comment removed",
            ChangeKind::CommentChanged => "comment changed",
            ChangeKind::Value => "value",
        }
    }
}

/// One cell's change, or a block of neighbouring cells that changed in the same way.
#[derive(Clone, Debug)]
pub struct CellChange {
    pub kind: ChangeKind,
    /// The cell or range in the old version, when it existed there.
    pub old: Option<String>,
    /// The cell or range in the new version, when it exists there.
    pub new: Option<String>,
    /// How many cells this entry covers.
    pub count: u32,
    pub before: String,
    pub after: String,
}

/// Compares `old` with `new`.
pub fn compare(old: &Book, new: &Book, options: Options) -> Report {
    let mut report = Report::default();

    // Tabs: by name, then renamed by content.
    let mut pairs: Vec<(usize, usize)> = Vec::new();
    let mut old_unmatched: Vec<usize> = Vec::new();
    let mut new_matched = vec![false; new.sheets.len()];
    for (i, s) in old.sheets.iter().enumerate() {
        match new.sheets.iter().position(|t| t.name == s.name) {
            Some(j) => {
                pairs.push((i, j));
                new_matched[j] = true;
            }
            None => old_unmatched.push(i),
        }
    }
    let mut new_unmatched: Vec<usize> =
        (0..new.sheets.len()).filter(|&j| !new_matched[j]).collect();
    let mut candidates: Vec<(f64, usize, usize)> = Vec::new();
    for &i in &old_unmatched {
        for &j in &new_unmatched {
            let s = sheet_similarity(&old.sheets[i], &new.sheets[j]);
            if s >= 0.5 {
                candidates.push((s, i, j));
            }
        }
    }
    candidates.sort_by(|a, b| b.0.total_cmp(&a.0));
    for (_, i, j) in candidates {
        if old_unmatched.contains(&i) && new_unmatched.contains(&j) {
            old_unmatched.retain(|&k| k != i);
            new_unmatched.retain(|&k| k != j);
            pairs.push((i, j));
            report
                .renamed_sheets
                .push((old.sheets[i].name.clone(), new.sheets[j].name.clone()));
        }
    }
    report.deleted_sheets = old_unmatched
        .iter()
        .map(|&i| old.sheets[i].name.clone())
        .collect();
    report.added_sheets = new_unmatched
        .iter()
        .map(|&j| new.sheets[j].name.clone())
        .collect();

    // Order, among the tabs in both versions.
    let mut by_old = pairs.clone();
    by_old.sort_by_key(|p| p.0);
    let old_order: Vec<String> = by_old
        .iter()
        .map(|&(_, j)| new.sheets[j].name.clone())
        .collect();
    let mut by_new = pairs.clone();
    by_new.sort_by_key(|p| p.1);
    let new_order: Vec<String> = by_new
        .iter()
        .map(|&(_, j)| new.sheets[j].name.clone())
        .collect();
    if old_order != new_order {
        let mut in_place = BTreeSet::new();
        for op in capture_diff_slices(Algorithm::Myers, &old_order, &new_order) {
            if let DiffOp::Equal { old_index, len, .. } = op {
                for name in &old_order[old_index..old_index + len] {
                    in_place.insert(name.clone());
                }
            }
        }
        for (pos, name) in new_order.iter().enumerate() {
            if !in_place.contains(name) {
                let old_pos = old_order.iter().position(|n| n == name).unwrap_or(0);
                report
                    .reordered_sheets
                    .push((name.clone(), old_pos + 1, pos + 1));
            }
        }
        report.order = Some((
            old.sheets.iter().map(|s| s.name.clone()).collect(),
            new.sheets.iter().map(|s| s.name.clone()).collect(),
        ));
    }

    // Pass 1: align every pair of tabs.
    let mut alignments: HashMap<String, (String, Alignment, Alignment)> = HashMap::new();
    for &(i, j) in &pairs {
        let (o, n) = (&old.sheets[i], &new.sheets[j]);
        let (rows, cols) = align_sheets(o, n);
        alignments.insert(o.name.clone(), (n.name.clone(), rows, cols));
    }

    // Pass 2: compare, rewriting old references through every tab's alignment.
    let deleted: BTreeSet<&str> = report.deleted_sheets.iter().map(String::as_str).collect();
    for &(i, j) in &by_new {
        let (o, n) = (&old.sheets[i], &new.sheets[j]);
        let lookup = |sheet: Option<&str>| -> Option<SheetMap<'_>> {
            let key = sheet.unwrap_or(&o.name);
            if deleted.contains(key) {
                return None;
            }
            alignments.get(key).map(|(new_name, rows, cols)| SheetMap {
                new_name,
                rows: rows as &dyn Remap,
                cols: cols as &dyn Remap,
            })
        };
        let (rows, cols) = {
            let (_, r, c) = &alignments[&o.name];
            (r, c)
        };
        let mut sr = compare_sheet(o, n, rows, cols, &lookup, options);
        if o.name != n.name {
            sr.renamed_from = Some(o.name.clone());
        }
        if sr.is_empty() {
            report.unchanged_sheets.push(n.name.clone());
        } else {
            report.sheets.push(sr);
        }
    }

    // Defined names.
    let lookup_any = |sheet: Option<&str>| -> Option<SheetMap<'_>> {
        let key = sheet?;
        if deleted.contains(key) {
            return None;
        }
        alignments.get(key).map(|(new_name, rows, cols)| SheetMap {
            new_name,
            rows: rows as &dyn Remap,
            cols: cols as &dyn Remap,
        })
    };
    let key = |scope: &Option<String>, name: &str| match scope {
        Some(s) => format!("{s}!{name}"),
        None => name.to_owned(),
    };
    let renamed: HashMap<&str, &str> = report
        .renamed_sheets
        .iter()
        .map(|(a, b)| (a.as_str(), b.as_str()))
        .collect();
    let old_names: BTreeMap<String, String> = old
        .names
        .iter()
        .map(|n| {
            let scope = n.scope.as_ref().map(|s| {
                renamed
                    .get(s.as_str())
                    .map_or(s.clone(), |r| (*r).to_owned())
            });
            (
                key(&scope, &n.name),
                formula::remap(&n.refers_to, &lookup_any),
            )
        })
        .collect();
    let new_names: BTreeMap<String, String> = new
        .names
        .iter()
        .map(|n| (key(&n.scope, &n.name), formula::normalized(&n.refers_to)))
        .collect();
    for (k, before) in &old_names {
        match new_names.get(k) {
            None => report.names.push(NameChange::Removed {
                name: k.clone(),
                refers_to: before.clone(),
            }),
            Some(after) if after != before => report.names.push(NameChange::Changed {
                name: k.clone(),
                before: before.clone(),
                after: after.clone(),
            }),
            Some(_) => {}
        }
    }
    for (k, after) in &new_names {
        if !old_names.contains_key(k) {
            report.names.push(NameChange::Added {
                name: k.clone(),
                refers_to: after.clone(),
            });
        }
    }
    report
}

fn hash_str(s: &str) -> u64 {
    let mut h = DefaultHasher::new();
    s.hash(&mut h);
    h.finish()
}

/// Content signatures of a sheet's cells, without positions.
fn content_bag(sheet: &Sheet) -> HashMap<String, usize> {
    let mut bag = HashMap::new();
    for (&(r, c), cell) in &sheet.cells {
        if cell.has_content() {
            *bag.entry(cell.content_signature(r, c)).or_insert(0) += 1;
        }
    }
    bag
}

fn bag_similarity(a: &HashMap<String, usize>, b: &HashMap<String, usize>) -> f64 {
    let total_a: usize = a.values().sum();
    let total_b: usize = b.values().sum();
    if total_a == 0 && total_b == 0 {
        return 1.0;
    }
    let common: usize = a
        .iter()
        .map(|(k, &n)| n.min(b.get(k).copied().unwrap_or(0)))
        .sum();
    common as f64 / total_a.max(total_b) as f64
}

fn sheet_similarity(a: &Sheet, b: &Sheet) -> f64 {
    bag_similarity(&content_bag(a), &content_bag(b))
}

/// Aligns columns by their content, then rows by their content in the aligned columns.
fn align_sheets(o: &Sheet, n: &Sheet) -> (Alignment, Alignment) {
    let blank = hash_str("");

    // Columns.
    let col_bags = |s: &Sheet| -> Vec<HashMap<String, usize>> {
        let mut bags = vec![HashMap::new(); s.max_col as usize];
        for (&(r, c), cell) in &s.cells {
            if cell.has_content() && c <= s.max_col {
                // Relative to the column only, so that row inserts do not disturb it.
                *bags[c as usize - 1]
                    .entry(cell.content_signature(r, c))
                    .or_insert(0) += 1;
            }
        }
        bags
    };
    let bag_sig = |bag: &HashMap<String, usize>| -> u64 {
        if bag.is_empty() {
            return blank;
        }
        let mut items: Vec<(&String, &usize)> = bag.iter().collect();
        items.sort();
        let mut h = DefaultHasher::new();
        items.hash(&mut h);
        h.finish()
    };
    let (ob, nb) = (col_bags(o), col_bags(n));
    let osig: Vec<u64> = ob.iter().map(bag_sig).collect();
    let nsig: Vec<u64> = nb.iter().map(bag_sig).collect();
    let cols = align(&osig, &nsig, blank, &|i, j| bag_similarity(&ob[i], &nb[j]));

    // Rows: each old cell's formula rewritten through the column alignment, then signed relative
    // to its new position, so that a column insert does not make every row look different.
    let col_pairs: Vec<(u32, u32)> = cols.pairs().collect();
    let identity = Alignment::default();
    let lookup = |sheet: Option<&str>| {
        sheet.is_none().then_some(SheetMap {
            new_name: "",
            rows: &identity as &dyn Remap,
            cols: &cols as &dyn Remap,
        })
    };
    let old_cell_sig = |cell: &Cell, r: u32, nc: u32| -> String {
        match &cell.content {
            Content::Formula { tokens, .. } => {
                let text = formula::remap(tokens, &lookup);
                format!(
                    "={}",
                    formula::relative_form(&formula::tokenize(&text), r, nc)
                )
            }
            _ => cell.content_signature(r, nc),
        }
    };
    let row_cells = |s: &Sheet, old_side: bool| -> Vec<Vec<String>> {
        let mut rows = vec![vec![String::new(); col_pairs.len()]; s.max_row as usize];
        for (k, &(oc, nc)) in col_pairs.iter().enumerate() {
            let c = if old_side { oc } else { nc };
            for r in 1..=s.max_row {
                if let Some(cell) = s.cells.get(&(r, c))
                    && cell.has_content()
                {
                    rows[r as usize - 1][k] = if old_side {
                        old_cell_sig(cell, r, nc)
                    } else {
                        cell.content_signature(r, nc)
                    };
                }
            }
        }
        rows
    };
    let (orows, nrows) = (row_cells(o, true), row_cells(n, false));
    let row_sig = |cells: &Vec<String>| -> u64 {
        if cells.iter().all(String::is_empty) {
            blank
        } else {
            let mut h = DefaultHasher::new();
            cells.hash(&mut h);
            h.finish()
        }
    };
    let osig: Vec<u64> = orows.iter().map(row_sig).collect();
    let nsig: Vec<u64> = nrows.iter().map(row_sig).collect();
    let labels = |s: &Sheet| -> Vec<Option<String>> {
        (1..=s.max_row)
            .map(|r| {
                s.cells
                    .range((r, 0)..=(r, u32::MAX))
                    .find(|(_, c)| c.has_content())
                    .and_then(|(_, c)| match &c.content {
                        Content::Value(Value::Text(t)) if !t.trim().is_empty() => Some(t.clone()),
                        _ => None,
                    })
            })
            .collect()
    };
    let (olabels, nlabels) = (labels(o), labels(n));
    let rows = align_labelled(
        &osig,
        &nsig,
        blank,
        &|i, j| {
            let (a, b) = (&orows[i], &nrows[j]);
            let mut used = 0;
            let mut same = 0;
            for k in 0..a.len() {
                if !a[k].is_empty() || !b[k].is_empty() {
                    used += 1;
                    if a[k] == b[k] {
                        same += 1;
                    }
                }
            }
            if used == 0 {
                return 1.0;
            }
            let fraction = same as f64 / used as f64;
            // A row's leading text is its label: a row that kept its label and changed every formula
            // is still the same row.
            let leading = |cells: &Vec<String>| cells.iter().position(|c| !c.is_empty());
            match (leading(a), leading(b)) {
                (Some(x), Some(y)) if x == y && a[x] == b[y] && a[x].starts_with('s') => {
                    fraction.max(0.75)
                }
                _ => fraction,
            }
        },
        Some((&olabels, &nlabels)),
    );
    (rows, cols)
}

/// An entry before grouping: one cell.
struct Raw {
    kind: ChangeKind,
    old: Option<(u32, u32)>,
    new: Option<(u32, u32)>,
    before: String,
    after: String,
    /// What must be equal between neighbours for them to be grouped.
    shape: String,
}

fn compare_sheet<'a>(
    o: &Sheet,
    n: &Sheet,
    rows: &Alignment,
    cols: &Alignment,
    lookup: &dyn Fn(Option<&str>) -> Option<SheetMap<'a>>,
    options: Options,
) -> SheetReport {
    let mut sr = SheetReport {
        name: n.name.clone(),
        ..SheetReport::default()
    };

    // Structure: rows and columns.
    let describe = |what: &str, side: &str, list: &[u32], letters: bool| -> Option<String> {
        if list.is_empty() {
            return None;
        }
        let fmt = |i: u32| {
            if letters {
                column_letters(i)
            } else {
                i.to_string()
            }
        };
        let runs = runs(list)
            .into_iter()
            .map(|(a, b)| {
                if a == b {
                    fmt(a)
                } else {
                    format!("{}–{}", fmt(a), fmt(b))
                }
            })
            .collect::<Vec<_>>()
            .join(", ");
        Some(format!("{} {what} {side}: {runs}", list.len()))
    };
    let unit = |n: usize, one: &str, many: &str| if n == 1 { one } else { many }.to_owned();
    if let Some(s) = describe(
        &unit(rows.inserted.len(), "row", "rows"),
        "inserted (new numbering)",
        &rows.inserted,
        false,
    ) {
        sr.structure.push(s);
    }
    if let Some(s) = describe(
        &unit(rows.deleted.len(), "row", "rows"),
        "deleted (old numbering)",
        &rows.deleted,
        false,
    ) {
        sr.structure.push(s);
    }
    if let Some(s) = describe(
        &unit(cols.inserted.len(), "column", "columns"),
        "inserted (new lettering)",
        &cols.inserted,
        true,
    ) {
        sr.structure.push(s);
    }
    if let Some(s) = describe(
        &unit(cols.deleted.len(), "column", "columns"),
        "deleted (old lettering)",
        &cols.deleted,
        true,
    ) {
        sr.structure.push(s);
    }
    for &(a, b) in &rows.moved {
        sr.structure.push(format!("Row {a} moved to row {b}"));
    }
    for &(a, b) in &cols.moved {
        sr.structure.push(format!(
            "Column {} moved to column {}",
            column_letters(a),
            column_letters(b)
        ));
    }

    // Cells.
    let mut raws: Vec<Raw> = Vec::new();
    let mut values: Vec<Raw> = Vec::new();
    let mut newly_unlocked: Vec<(u32, u32)> = Vec::new();
    let mut newly_locked: Vec<(u32, u32)> = Vec::new();
    let mut restyled: Vec<(u32, u32)> = Vec::new();
    let row_pairs: Vec<(u32, u32)> = rows.pairs().collect();
    let col_pairs: Vec<(u32, u32)> = cols.pairs().collect();
    let empty = Cell {
        content: Content::Empty,
        cached: None,
        comment: None,
        unlocked: false,
        style: None,
    };
    for &(or, nr) in &row_pairs {
        for &(oc, nc) in &col_pairs {
            let a = o.cells.get(&(or, oc));
            let b = n.cells.get(&(nr, nc));
            if a.is_none() && b.is_none() {
                continue;
            }
            let (a, b) = (a.unwrap_or(&empty), b.unwrap_or(&empty));
            compare_cell(a, b, (or, oc), (nr, nc), lookup, &mut raws, &mut values);
            if a.unlocked != b.unlocked {
                if b.unlocked {
                    newly_unlocked.push((nr, nc));
                } else {
                    newly_locked.push((nr, nc));
                }
            }
            if options.formats && a.style.is_some() && b.style.is_some() && a.style != b.style {
                restyled.push((nr, nc));
            }
        }
    }
    // Cells in inserted rows or columns.
    let inserted_rows: BTreeSet<u32> = rows.inserted.iter().copied().collect();
    let inserted_cols: BTreeSet<u32> = cols.inserted.iter().copied().collect();
    for (&(r, c), cell) in &n.cells {
        if inserted_rows.contains(&r) || inserted_cols.contains(&c) {
            if cell.has_content() {
                raws.push(Raw {
                    kind: ChangeKind::Added,
                    old: None,
                    new: Some((r, c)),
                    before: String::new(),
                    after: describe_content(cell),
                    shape: cell.content_signature(r, c),
                });
            }
            if let Some(text) = &cell.comment {
                raws.push(Raw {
                    kind: ChangeKind::CommentAdded,
                    old: None,
                    new: Some((r, c)),
                    before: String::new(),
                    after: text.clone(),
                    shape: text.clone(),
                });
            }
            if cell.unlocked {
                newly_unlocked.push((r, c));
            }
        }
    }
    let deleted_rows: BTreeSet<u32> = rows.deleted.iter().copied().collect();
    let deleted_cols: BTreeSet<u32> = cols.deleted.iter().copied().collect();
    for (&(r, c), cell) in &o.cells {
        if deleted_rows.contains(&r) || deleted_cols.contains(&c) {
            if cell.has_content() {
                raws.push(Raw {
                    kind: ChangeKind::Removed,
                    old: Some((r, c)),
                    new: None,
                    before: describe_content(cell),
                    after: String::new(),
                    shape: cell.content_signature(r, c),
                });
            }
            if let Some(text) = &cell.comment {
                raws.push(Raw {
                    kind: ChangeKind::CommentRemoved,
                    old: Some((r, c)),
                    new: None,
                    before: text.clone(),
                    after: String::new(),
                    shape: text.clone(),
                });
            }
        }
    }
    sr.cells = group(raws);
    if options.values {
        sr.values = group(values);
    }

    // Unlocked (input) cells.
    if !newly_unlocked.is_empty() {
        sr.structure.push(format!(
            "Cells unlocked: {}",
            ranges(&newly_unlocked).join(", ")
        ));
    }
    if !newly_locked.is_empty() {
        sr.structure.push(format!(
            "Cells locked: {}",
            ranges(&newly_locked).join(", ")
        ));
    }

    // Merged cells, conditional formats and validations, compared in new coordinates.
    let remap_range = |range: &str| formula::remap(&formula::tokenize(range), lookup);
    let set_diff = |sr: &mut SheetReport, what: &str, before: Vec<String>, after: Vec<String>| {
        let mut b = before.clone();
        for x in &after {
            if let Some(p) = b.iter().position(|y| y == x) {
                b.remove(p);
            } else {
                sr.structure.push(format!("{what} added: {x}"));
            }
        }
        for x in b {
            sr.structure.push(format!("{what} removed: {x}"));
        }
    };
    set_diff(
        &mut sr,
        "Merged cells",
        o.merges
            .iter()
            .map(|m| format!("`{}`", remap_range(m)))
            .collect(),
        n.merges.iter().map(|m| format!("`{m}`")).collect(),
    );
    let cf_text = |ranges: Vec<String>, rule: &str, f: Option<String>| {
        let mut s = format!("`{}` {rule}", ranges.join(" "));
        if let Some(f) = f {
            s.push_str(&format!(" `={f}`"));
        }
        s
    };
    set_diff(
        &mut sr,
        "Conditional format",
        o.conditional_formats
            .iter()
            .map(|cf| {
                cf_text(
                    cf.ranges.iter().map(|r| remap_range(r)).collect(),
                    &cf.rule,
                    cf.formula.as_ref().map(|t| formula::remap(t, lookup)),
                )
            })
            .collect(),
        n.conditional_formats
            .iter()
            .map(|cf| {
                cf_text(
                    cf.ranges.clone(),
                    &cf.rule,
                    cf.formula.as_ref().map(|t| formula::normalized(t)),
                )
            })
            .collect(),
    );
    let dv_text = |ranges: Vec<String>, rule: &str, f1: String, f2: String| {
        let mut s = format!("`{}` {rule}", ranges.join(" "));
        if !f1.is_empty() {
            s.push_str(&format!(" `{f1}`"));
        }
        if !f2.is_empty() {
            s.push_str(&format!(" `{f2}`"));
        }
        s
    };
    set_diff(
        &mut sr,
        "Data validation",
        o.validations
            .iter()
            .map(|v| {
                dv_text(
                    v.ranges.iter().map(|r| remap_range(r)).collect(),
                    &v.rule,
                    formula::remap(&v.formula1, lookup),
                    formula::remap(&v.formula2, lookup),
                )
            })
            .collect(),
        n.validations
            .iter()
            .map(|v| {
                dv_text(
                    v.ranges.clone(),
                    &v.rule,
                    formula::normalized(&v.formula1),
                    formula::normalized(&v.formula2),
                )
            })
            .collect(),
    );

    // Protection, visibility and panes.
    if o.protection != n.protection {
        let say = |p: &Option<super::model::Protection>| match p {
            None => "not protected".to_owned(),
            Some(p) if p.has_password => "protected with a password".to_owned(),
            Some(_) => "protected without a password".to_owned(),
        };
        sr.structure.push(format!(
            "Sheet protection: {} → {}",
            say(&o.protection),
            say(&n.protection)
        ));
    }
    if o.state != n.state {
        sr.structure
            .push(format!("Tab visibility: {} → {}", o.state, n.state));
    }
    if o.frozen != n.frozen {
        let say = |f: Option<(u32, u32)>| match f {
            None => "none".to_owned(),
            Some((r, c)) => format!("{r} rows, {c} columns"),
        };
        sr.structure.push(format!(
            "Frozen panes: {} → {}",
            say(o.frozen),
            say(n.frozen)
        ));
    }
    let hidden_diff = |sr: &mut SheetReport,
                       what: &str,
                       old: &BTreeSet<u32>,
                       new: &BTreeSet<u32>,
                       map: &Alignment,
                       letters: bool| {
        let mapped: BTreeSet<u32> = old.iter().filter_map(|&i| map.exact(i)).collect();
        let fmt = |v: Vec<u32>| {
            runs(&v)
                .into_iter()
                .map(|(a, b)| {
                    let f = |i| {
                        if letters {
                            column_letters(i)
                        } else {
                            i.to_string()
                        }
                    };
                    if a == b {
                        f(a)
                    } else {
                        format!("{}–{}", f(a), f(b))
                    }
                })
                .collect::<Vec<_>>()
                .join(", ")
        };
        let hidden: Vec<u32> = new.difference(&mapped).copied().collect();
        let shown: Vec<u32> = mapped.difference(new).copied().collect();
        if !hidden.is_empty() {
            sr.structure.push(format!("{what} hidden: {}", fmt(hidden)));
        }
        if !shown.is_empty() {
            sr.structure
                .push(format!("{what} unhidden: {}", fmt(shown)));
        }
    };
    hidden_diff(&mut sr, "Rows", &o.hidden_rows, &n.hidden_rows, rows, false);
    hidden_diff(
        &mut sr,
        "Columns",
        &o.hidden_cols,
        &n.hidden_cols,
        cols,
        true,
    );

    // Formats.
    if options.formats {
        if !restyled.is_empty() {
            sr.formats.push(format!(
                "Cell formats changed: {}",
                ranges(&restyled).join(", ")
            ));
        }
        let mut widths = Vec::new();
        for &(oc, nc) in &col_pairs {
            let (a, b) = (o.col_widths.get(&oc), n.col_widths.get(&nc));
            // A resave rounds widths to the application's own units, moving them by a few
            // hundredths of a character; a tenth is the smallest change anyone would see.
            if let (Some(a), Some(b)) = (a, b)
                && (a - b).abs() >= 0.1
            {
                widths.push(format!("{} {a:.2} → {b:.2}", column_letters(nc)));
            }
        }
        if !widths.is_empty() {
            sr.formats
                .push(format!("Column widths: {}", widths.join(", ")));
        }
        let mut heights = Vec::new();
        for &(or, nr) in &row_pairs {
            let (a, b) = (o.row_heights.get(&or), n.row_heights.get(&nr));
            if a.is_some() != b.is_some() || a.zip(b).is_some_and(|(a, b)| (a - b).abs() >= 0.6) {
                let say = |h: Option<&f64>| h.map_or("default".to_owned(), |h| format!("{h:.2}"));
                heights.push(format!("{nr} {} → {}", say(a), say(b)));
            }
        }
        if !heights.is_empty() {
            sr.formats
                .push(format!("Row heights: {}", heights.join(", ")));
        }
    }
    sr
}

#[allow(clippy::too_many_arguments)]
fn compare_cell<'a>(
    a: &Cell,
    b: &Cell,
    at_old: (u32, u32),
    at_new: (u32, u32),
    lookup: &dyn Fn(Option<&str>) -> Option<SheetMap<'a>>,
    raws: &mut Vec<Raw>,
    values: &mut Vec<Raw>,
) {
    let (nr, nc) = at_new;
    let content_changed = match (&a.content, &b.content) {
        (Content::Empty, Content::Empty) | (Content::ArrayMember, Content::ArrayMember) => None,
        (Content::Value(x), Content::Value(y)) => (!x.same(y)).then(String::new),
        (
            Content::Formula {
                tokens: ta,
                array: aa,
                ..
            },
            Content::Formula {
                tokens: tb,
                array: ab,
                ..
            },
        ) => {
            let before = formula::remap(ta, lookup);
            let after = formula::normalized(tb);
            let shape = format!(
                "{}|{}",
                formula::relative_form(&formula::tokenize(&before), nr, nc),
                formula::relative_form(tb, nr, nc)
            );
            (before != after || aa.is_some() != ab.is_some()).then_some(shape)
        }
        _ => Some(String::new()),
    };
    if let Some(mut shape) = content_changed {
        if shape.is_empty() {
            shape = format!(
                "{}|{}",
                a.content_signature(at_old.0, at_old.1),
                b.content_signature(nr, nc)
            );
        }
        raws.push(Raw {
            kind: ChangeKind::Changed,
            old: Some(at_old),
            new: Some(at_new),
            before: describe_content(a),
            after: describe_content(b),
            shape,
        });
    } else if let (Some(x), Some(y)) = (&a.cached, &b.cached)
        && !x.same(y)
    {
        values.push(Raw {
            kind: ChangeKind::Value,
            old: Some(at_old),
            new: Some(at_new),
            before: x.display(),
            after: y.display(),
            shape: format!("{}|{}", x.display(), y.display()),
        });
    }
    match (&a.comment, &b.comment) {
        (None, Some(t)) => raws.push(Raw {
            kind: ChangeKind::CommentAdded,
            old: Some(at_old),
            new: Some(at_new),
            before: String::new(),
            after: t.clone(),
            shape: t.clone(),
        }),
        (Some(t), None) => raws.push(Raw {
            kind: ChangeKind::CommentRemoved,
            old: Some(at_old),
            new: Some(at_new),
            before: t.clone(),
            after: String::new(),
            shape: t.clone(),
        }),
        (Some(x), Some(y)) if x != y => raws.push(Raw {
            kind: ChangeKind::CommentChanged,
            old: Some(at_old),
            new: Some(at_new),
            before: x.clone(),
            after: y.clone(),
            shape: format!("{x}|{y}"),
        }),
        _ => {}
    }
}

fn describe_content(cell: &Cell) -> String {
    match &cell.content {
        Content::Empty => String::new(),
        Content::Value(Value::Text(t)) => format!("\"{t}\""),
        Content::Value(v) => v.display(),
        Content::Formula { text, array, .. } => match array {
            Some(r) => format!("{{={text}}} over {r}"),
            None => format!("={text}"),
        },
        Content::ArrayMember => "(part of an array formula)".to_owned(),
    }
}

/// A rectangle of cells: (first row, first column, last row, last column).
type Rect = (u32, u32, u32, u32);

fn rect_name(r: Option<Rect>) -> Option<String> {
    r.map(|(r1, c1, r2, c2)| {
        if (r1, c1) == (r2, c2) {
            cell_name(r1, c1)
        } else {
            format!("{}:{}", cell_name(r1, c1), cell_name(r2, c2))
        }
    })
}

/// What joins two runs across a row: kind, shape, and the old and new (first row, last row,
/// column) at which one ends or the other must start.
type RunKey = (u8, String, Option<(u32, u32, u32)>, Option<(u32, u32, u32)>);

/// A run of cells changed in the same way, while grouping.
struct Run {
    raw: Raw,
    old: Option<Rect>,
    new: Option<Rect>,
    count: u32,
}

/// Groups single-cell entries changed in the same way: first into runs down a column, then runs
/// of the same rows in adjacent columns into one rectangle, so that a change copied across a
/// row is one entry too.
fn group(mut raws: Vec<Raw>) -> Vec<CellChange> {
    let pos = |r: &Raw| r.new.or(r.old).unwrap_or((0, 0));
    raws.sort_by_key(|r| {
        let (row, col) = pos(r);
        (col, row, r.kind as u8)
    });
    let cell = |at: Option<(u32, u32)>| at.map(|(r, c)| (r, c, r, c));

    // Down each column.
    let mut runs: Vec<Run> = Vec::new();
    for raw in raws {
        if let Some(prev) = runs.last_mut() {
            let follows = |range: Option<Rect>, at: Option<(u32, u32)>| match (range, at) {
                (Some((_, c1, r2, c2)), Some((r, c))) => c1 == c2 && c == c2 && r == r2 + 1,
                (None, None) => true,
                _ => false,
            };
            if prev.raw.kind == raw.kind
                && prev.raw.shape == raw.shape
                && follows(prev.old, raw.old)
                && follows(prev.new, raw.new)
            {
                prev.count += 1;
                if let (Some(r), Some((row, _))) = (prev.old.as_mut(), raw.old) {
                    r.2 = row;
                }
                if let (Some(r), Some((row, _))) = (prev.new.as_mut(), raw.new) {
                    r.2 = row;
                }
                continue;
            }
        }
        runs.push(Run {
            old: cell(raw.old),
            new: cell(raw.new),
            count: 1,
            raw,
        });
    }

    // Across a row: a run joins the run in the column to its left with the same rows.
    let first = |r: &Run| r.new.or(r.old).unwrap_or_default();
    runs.sort_by_key(|r| {
        let (r1, c1, r2, _) = first(r);
        (r1, r2, c1, r.raw.kind as u8)
    });
    let mut out: Vec<Run> = Vec::new();
    // Where each group ends, so the run starting in the next column can find it.
    let mut open: HashMap<RunKey, usize> = HashMap::new();
    let ends = |r: Option<Rect>| r.map(|(r1, _, r2, c2)| (r1, r2, c2));
    let before = |r: Option<Rect>| r.map(|(r1, c1, r2, _)| (r1, r2, c1.wrapping_sub(1)));
    for run in runs {
        let key = (
            run.raw.kind as u8,
            run.raw.shape.clone(),
            before(run.old),
            before(run.new),
        );
        if let Some(i) = open.remove(&key) {
            let g = &mut out[i];
            g.count += run.count;
            if let (Some(r), Some(n)) = (g.old.as_mut(), run.old) {
                r.3 = n.3;
            }
            if let (Some(r), Some(n)) = (g.new.as_mut(), run.new) {
                r.3 = n.3;
            }
            let k = (
                g.raw.kind as u8,
                g.raw.shape.clone(),
                ends(g.old),
                ends(g.new),
            );
            open.insert(k, i);
            continue;
        }
        let k = (
            run.raw.kind as u8,
            run.raw.shape.clone(),
            ends(run.old),
            ends(run.new),
        );
        open.insert(k, out.len());
        out.push(run);
    }
    out.sort_by_key(|r| {
        let (r1, c1, _, _) = first(r);
        (c1, r1, r.raw.kind as u8)
    });
    out.into_iter()
        .map(|r| CellChange {
            kind: r.raw.kind,
            old: rect_name(r.old),
            new: rect_name(r.new),
            count: r.count,
            before: r.raw.before,
            after: r.raw.after,
        })
        .collect()
}

/// Consecutive runs in an ascending list.
fn runs(list: &[u32]) -> Vec<(u32, u32)> {
    let mut sorted = list.to_vec();
    sorted.sort_unstable();
    sorted.dedup();
    let mut out: Vec<(u32, u32)> = Vec::new();
    for i in sorted {
        match out.last_mut() {
            Some((_, b)) if *b + 1 == i => *b = i,
            _ => out.push((i, i)),
        }
    }
    out
}

/// Compact rectangular ranges covering a set of cells: runs down each column, merged across
/// adjacent columns with the same run.
fn ranges(cells: &[(u32, u32)]) -> Vec<String> {
    let mut by_col: BTreeMap<u32, Vec<u32>> = BTreeMap::new();
    for &(r, c) in cells {
        by_col.entry(c).or_default().push(r);
    }
    let mut blocks: Vec<(u32, u32, u32, u32)> = Vec::new(); // r1, r2, c1, c2
    for (c, rows) in by_col {
        for (r1, r2) in runs(&rows) {
            if let Some(b) = blocks
                .iter_mut()
                .find(|b| b.0 == r1 && b.1 == r2 && b.3 + 1 == c)
            {
                b.3 = c;
            } else {
                blocks.push((r1, r2, c, c));
            }
        }
    }
    blocks.sort_by_key(|b| (b.2, b.0));
    blocks
        .into_iter()
        .map(|(r1, r2, c1, c2)| {
            if r1 == r2 && c1 == c2 {
                cell_name(r1, c1)
            } else {
                format!("{}:{}", cell_name(r1, c1), cell_name(r2, c2))
            }
        })
        .collect()
}
