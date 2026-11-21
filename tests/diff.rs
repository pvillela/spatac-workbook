//! `xlsx-diff` against small workbooks generated here.

use spatac_workbook::diff::compare::{ChangeKind, NameChange, SheetReport};
use spatac_workbook::diff::{Options, Report, diff_files, markdown};
use std::path::PathBuf;
use umya_spreadsheet::Workbook;
use umya_spreadsheet::structs::{
    ConditionalFormatValues, ConditionalFormatting, ConditionalFormattingRule, Formula,
    SequenceOfReferences,
};

/// A fresh directory for one test's files.
fn dir(test: &str) -> PathBuf {
    let d = std::env::temp_dir().join(format!("xlsx-diff-tests-{}-{test}", std::process::id()));
    let _ = std::fs::remove_dir_all(&d);
    std::fs::create_dir_all(&d).unwrap();
    d
}

/// A workbook with a sheet `Data`: A1:A5 = 1..5, B1:B5 = A*10 as formulas, B6 = SUM(B1:B5).
fn base() -> Workbook {
    let mut wb = umya_spreadsheet::new_file();
    let ws = wb.sheet_mut(0).unwrap();
    ws.set_name("Data");
    for r in 1..=5u32 {
        ws.cell_mut((1, r)).set_value_number(r as f64);
        ws.cell_mut((2, r)).set_formula(format!("A{r}*10"));
    }
    ws.cell_mut("B6").set_formula("SUM(B1:B5)");
    wb
}

fn run(test: &str, old: &Workbook, new: &Workbook, options: Options) -> Report {
    let d = dir(test);
    let (a, b) = (d.join("old.xlsx"), d.join("new.xlsx"));
    umya_spreadsheet::writer::xlsx::write(old, &a).unwrap();
    umya_spreadsheet::writer::xlsx::write(new, &b).unwrap();
    diff_files(&a, &b, options).unwrap()
}

fn sheet<'a>(r: &'a Report, name: &str) -> &'a SheetReport {
    r.sheets
        .iter()
        .find(|s| s.name == name)
        .unwrap_or_else(|| panic!("no report for {name}: {r:#?}"))
}

#[test]
fn identical_workbooks_have_no_differences() {
    let r = run("identical", &base(), &base(), Options::default());
    assert!(r.is_empty(), "{r:#?}");
    assert!(markdown::render(&r, "a", "b").contains("No differences."));
}

#[test]
fn inserted_row_is_structure_not_formula_changes() {
    let mut new = base();
    let ws = new.sheet_mut(0).unwrap();
    ws.insert_new_row(3, 1);
    ws.cell_mut("A3").set_value("inserted");
    let r = run("insert", &base(), &new, Options::default());
    let s = sheet(&r, "Data");
    assert!(
        s.structure
            .iter()
            .any(|t| t.contains("1 row inserted") && t.ends_with(": 3")),
        "{s:#?}"
    );
    // Only the new cell: the formulas below it moved, and B7's SUM grew with the insert.
    assert_eq!(s.cells.len(), 1, "{s:#?}");
    assert_eq!(s.cells[0].kind, ChangeKind::Added);
    assert_eq!(s.cells[0].new.as_deref(), Some("A3"));
}

#[test]
fn constant_and_comment_changes() {
    let mut new = base();
    let ws = new.sheet_mut(0).unwrap();
    ws.cell_mut("A2").set_value_number(20.0);
    let mut c = umya_spreadsheet::Comment::default();
    c.new_comment("A2");
    c.set_text_string("Revised");
    ws.add_comments(c);
    let r = run("constant", &base(), &new, Options::default());
    let s = sheet(&r, "Data");
    let kinds: Vec<ChangeKind> = s.cells.iter().map(|c| c.kind).collect();
    assert!(kinds.contains(&ChangeKind::Changed), "{s:#?}");
    assert!(kinds.contains(&ChangeKind::CommentAdded), "{s:#?}");
    let changed = s
        .cells
        .iter()
        .find(|c| c.kind == ChangeKind::Changed)
        .unwrap();
    assert_eq!(
        (changed.before.as_str(), changed.after.as_str()),
        ("2", "20")
    );
}

#[test]
fn filled_formula_change_is_one_entry() {
    let mut new = base();
    let ws = new.sheet_mut(0).unwrap();
    for r in 1..=5u32 {
        ws.cell_mut((2, r)).set_formula(format!("A{r}*100"));
    }
    let r = run("fill", &base(), &new, Options::default());
    let s = sheet(&r, "Data");
    assert_eq!(s.cells.len(), 1, "{s:#?}");
    assert_eq!(s.cells[0].count, 5);
    assert_eq!(s.cells[0].new.as_deref(), Some("B1:B5"));
}

#[test]
fn cross_sheet_reference_follows_an_insert_in_the_other_sheet() {
    let mut old = base();
    let other = old.new_sheet("Report").unwrap();
    other.cell_mut("A1").set_formula("Data!B6");
    let mut new = old.clone();
    let ws = new.sheet_by_name_mut("Data").unwrap();
    ws.insert_new_row(1, 2);
    // umya rewrites references on its own sheet; make the other sheet's follow, as an
    // application would.
    new.sheet_by_name_mut("Report")
        .unwrap()
        .cell_mut("A1")
        .set_formula("Data!B8");
    let r = run("cross", &old, &new, Options::default());
    assert!(
        r.sheets.iter().all(|s| s.name != "Report"),
        "Report should be unchanged: {r:#?}"
    );
}

#[test]
fn tabs_added_deleted_renamed_and_reordered() {
    let mut old = base();
    old.new_sheet("Gone").unwrap().cell_mut("A1").set_value("x");
    let summary = old.new_sheet("Summary").unwrap();
    for r in 1..=10u32 {
        summary.cell_mut((1, r)).set_value(format!("line {r}"));
    }
    let mut new = base();
    let renamed = new.new_sheet("Overview").unwrap();
    for r in 1..=10u32 {
        renamed.cell_mut((1, r)).set_value(format!("line {r}"));
    }
    new.new_sheet("Fresh")
        .unwrap()
        .cell_mut("A1")
        .set_value("y");
    // Move Data to the end.
    let data = new.sheet_collection()[0].clone();
    new.remove_sheet(0).unwrap();
    new.add_sheet(data).unwrap();
    let r = run("tabs", &old, &new, Options::default());
    assert_eq!(r.added_sheets, vec!["Fresh"]);
    assert_eq!(r.deleted_sheets, vec!["Gone"]);
    assert_eq!(
        r.renamed_sheets,
        vec![("Summary".to_owned(), "Overview".to_owned())]
    );
    assert!(r.order.is_some());
    assert!(!r.reordered_sheets.is_empty(), "{r:#?}");
}

#[test]
fn structural_items_are_reported() {
    let mut new = base();
    new.sheet_mut(0)
        .unwrap()
        .add_defined_name("Inputs", "Data!$A$1:$A$5")
        .unwrap();
    let ws = new.sheet_mut(0).unwrap();
    ws.style_mut("A1").protection_mut().set_locked(false);
    ws.sheet_protection_mut().set_sheet(true);
    let mut rule = ConditionalFormattingRule::default();
    rule.set_type(ConditionalFormatValues::Expression);
    let mut f = Formula::default();
    f.set_address_str("$A1>2");
    rule.set_formula(f);
    let mut seq = SequenceOfReferences::default();
    seq.set_sqref("A1:A5");
    let mut cf = ConditionalFormatting::default();
    cf.set_sequence_of_references(seq);
    cf.add_conditional_collection(rule);
    ws.add_conditional_formatting_collection(cf);
    ws.add_merge_cells("C1:D1");

    let r = run("structure", &base(), &new, Options::default());
    assert!(
        r.names
            .iter()
            .any(|n| matches!(n, NameChange::Added { name, .. } if name == "Inputs")),
        "{r:#?}"
    );
    let s = sheet(&r, "Data");
    let all = s.structure.join("\n");
    for expected in [
        "Cells unlocked: A1",
        "Sheet protection: not protected → protected without a password",
        "Conditional format added",
        "Merged cells added: `C1:D1`",
    ] {
        assert!(all.contains(expected), "missing {expected:?} in:\n{all}");
    }
}

#[test]
fn values_only_with_the_option() {
    let mut old = base();
    old.sheet_mut(0)
        .unwrap()
        .cell_mut("B1")
        .set_formula_result_default("10");
    let mut new = base();
    new.sheet_mut(0)
        .unwrap()
        .cell_mut("B1")
        .set_formula_result_default("99");
    let quiet = run("values-off", &old, &new, Options::default());
    assert!(quiet.is_empty(), "{quiet:#?}");
    let loud = run(
        "values-on",
        &old,
        &new,
        Options {
            values: true,
            ..Options::default()
        },
    );
    assert_eq!(sheet(&loud, "Data").values.len(), 1, "{loud:#?}");
}
