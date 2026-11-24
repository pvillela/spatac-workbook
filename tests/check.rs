//! `spatac-check` against small workbooks generated here.

use spatac_workbook::check::check;
use umya_spreadsheet::Workbook;
use umya_spreadsheet::structs::DefinedName;

fn name(wb: &mut Workbook, n: &str, address: &str) {
    let mut dn = DefinedName::default();
    dn.set_name(n);
    dn.set_address(address);
    wb.add_defined_names(dn);
}

/// Expenditures with a key column G (header row 1, items 2-4, ▼ End row 5); Adjustments with
/// one block: Item in A (inputs in rows 2-3) and its check in H.
fn workbook() -> Workbook {
    let mut wb = umya_spreadsheet::new_file();
    wb.sheet_mut(0).unwrap().set_name("Expenditures");
    wb.new_sheet("Adjustments").unwrap();
    {
        let ws = wb.sheet_by_name_mut("Expenditures").unwrap();
        ws.cell_mut("G1").set_value("Key");
        ws.cell_mut("G2")
            .set_formula("B2&C2")
            .set_formula_result_default("a");
        // G3 has lost its formula.
        ws.cell_mut("G3").set_value("typed over");
        ws.cell_mut("G4")
            .set_formula("B4&C4")
            .set_formula_result_default("#N/A");
        ws.cell_mut("G4").cell_value_mut().set_error("#N/A");
        ws.sheet_protection_mut().set_sheet(true);
    }
    {
        let ws = wb.sheet_by_name_mut("Adjustments").unwrap();
        for r in 2..=3 {
            for c in ["A", "B", "C", "D", "E"] {
                ws.style_mut(format!("{c}{r}").as_str())
                    .protection_mut()
                    .set_locked(false);
            }
        }
        ws.cell_mut("H2")
            .set_formula("A2")
            .set_formula_result_default("");
        // H3 is missing, and H2 is unlocked.
        ws.style_mut("H2").protection_mut().set_locked(false);
        // Adjustments is left unprotected.
    }
    name(&mut wb, "Exp_Key", "Expenditures!$G$1:$G$5");
    name(&mut wb, "Adj_Item", "Adjustments!$A$1:$A$4");
    name(&mut wb, "Adj_Check", "Adjustments!$H$1:$H$4");
    name(&mut wb, "Exp_Gone", "Expenditures!#REF!");
    wb
}

fn problems(test: &str) -> Vec<String> {
    let d = std::env::temp_dir().join(format!("check-tests-{}-{test}", std::process::id()));
    let _ = std::fs::remove_dir_all(&d);
    std::fs::create_dir_all(&d).unwrap();
    let path = d.join("Spatac-20261124.xlsx");
    umya_spreadsheet::writer::xlsx::write(&workbook(), &path).unwrap();
    check(&path)
        .unwrap()
        .iter()
        .map(|p| p.to_string())
        .collect()
}

fn has(problems: &[String], text: &str) -> bool {
    problems.iter().any(|p| p == text)
}

#[test]
fn every_problem_is_reported() {
    let p = problems("every");
    for expected in [
        "Expenditures!G3: a formula is missing (copy it from the row above)",
        "Expenditures!G4: the cell shows the error #N/A",
        "Adjustments!H3: a formula is missing (copy it from the row above)",
        "Adjustments!H2: a formula cell is unlocked",
        "Adjustments!B2: the input has lost its validation (copy the row above)",
        "Adjustments: the tab is not protected",
        "Exp_Gone: the name points at #REF!",
        "Cfg_BaseYear: the name is missing",
    ] {
        assert!(has(&p, expected), "{expected} not in {p:#?}");
    }
}

#[test]
fn good_cells_are_not_reported() {
    let p = problems("good");
    for clean in [
        "Expenditures!G2",
        "Expenditures!G5",
        "Adjustments!H1",
        "Adjustments!H4",
    ] {
        assert!(
            !p.iter().any(|x| x.starts_with(&format!("{clean}:"))),
            "{clean} reported: {p:#?}"
        );
    }
    assert!(!p.iter().any(|x| x.starts_with("Expenditures:")));
}

#[test]
fn problems_are_in_sheet_and_row_order() {
    let p = problems("order");
    let adj: Vec<&String> = p.iter().filter(|x| x.starts_with("Adjustments!")).collect();
    let rows: Vec<u32> = adj
        .iter()
        .map(|x| {
            let cell = &x["Adjustments!".len()..x.find(':').unwrap()];
            cell.trim_start_matches(char::is_alphabetic)
                .parse()
                .unwrap()
        })
        .collect();
    assert!(rows.windows(2).all(|w| w[0] <= w[1]), "{adj:#?}");
}
