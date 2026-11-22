//! `spatac-investment-plan` against a small workbook shaped like Spatac's, generated here.

use spatac_workbook::investment_plan::extract;
use std::path::PathBuf;
use umya_spreadsheet::Workbook;
use umya_spreadsheet::structs::DefinedName;

fn name(wb: &mut Workbook, n: &str, address: &str) {
    let mut dn = DefinedName::default();
    dn.set_name(n);
    dn.set_address(address);
    wb.add_defined_names(dn);
}

/// A formula with its cached result, as a saved workbook carries it.
fn formula(wb: &mut Workbook, sheet: &str, at: &str, f: &str, cached: &str) {
    let ws = wb.sheet_by_name_mut(sheet).unwrap();
    ws.cell_mut(at)
        .set_formula(f)
        .set_formula_result_default(cached);
}

fn spatac_shaped() -> Workbook {
    let mut wb = umya_spreadsheet::new_file();
    wb.sheet_mut(0).unwrap().set_name("Settings");
    for s in [
        "Expenditures",
        "Medium_Term_Cashflow",
        "Schedule_of_Investments",
    ] {
        wb.new_sheet(s).unwrap();
    }
    {
        let ws = wb.sheet_by_name_mut("Settings").unwrap();
        ws.cell_mut("B1").set_value_number(2026);
        ws.cell_mut("B2").set_value_number(0.03);
    }
    {
        // Items in A, elements in B, GL names in C, reasons in D, RFS 2026-28 in E:G, Revised in
        // H:J. Item 1.1 moves 100 from 2027 to 2028; item 1.2 is unchanged.
        let ws = wb.sheet_by_name_mut("Expenditures").unwrap();
        for (k, y) in [2026, 2027, 2028].iter().enumerate() {
            ws.cell_mut((5 + k as u32, 1)).set_value_number(*y);
            ws.cell_mut((8 + k as u32, 1)).set_value_number(*y);
        }
        ws.cell_mut("A2").set_value("1.1");
        ws.cell_mut("B2").set_value("Roof");
        ws.cell_mut("D2").set_value("Deferred one year");
        ws.cell_mut("F2").set_value_number(100);
        ws.cell_mut("J2").set_value_number(100);
        ws.cell_mut("A3").set_value("1.2");
        ws.cell_mut("B3").set_value("Doors");
        ws.cell_mut("E3").set_value_number(50);
        ws.cell_mut("H3").set_value_number(50);
    }
    {
        let ws = wb.sheet_by_name_mut("Medium_Term_Cashflow").unwrap();
        ws.cell_mut("A1").set_value("MEDIUM-TERM CASH FLOW");
        ws.cell_mut("B5").set_value_number(10);
        ws.cell_mut("A20")
            .set_value("Investment balance comparison (dropped)");
        ws.cell_mut("B20").set_value_number(1);
        ws.style_mut("B5").protection_mut().set_locked(false);
    }
    formula(
        &mut wb,
        "Medium_Term_Cashflow",
        "B3",
        "Cfg_BaseYear",
        "2026",
    );
    formula(&mut wb, "Medium_Term_Cashflow", "C3", "B3+1", "2027");
    formula(&mut wb, "Medium_Term_Cashflow", "D3", "C3+1", "2028");
    formula(&mut wb, "Medium_Term_Cashflow", "C5", "B5*2", "20");
    formula(
        &mut wb,
        "Medium_Term_Cashflow",
        "B6",
        "-SUM(Expenditures!H2:H3)",
        "-50",
    );
    formula(
        &mut wb,
        "Medium_Term_Cashflow",
        "B8",
        "SUM(Schedule_of_Investments!D3:D3)",
        "1000",
    );
    {
        let ws = wb.sheet_by_name_mut("Schedule_of_Investments").unwrap();
        ws.cell_mut("A1").set_value("SCHEDULE OF INVESTMENTS");
        ws.cell_mut("A3").set_value("GIC");
        ws.cell_mut("D3").set_value_number(1000);
    }
    formula(
        &mut wb,
        "Schedule_of_Investments",
        "B3",
        "D3*Cfg_NewInvRate",
        "30",
    );
    formula(&mut wb, "Schedule_of_Investments", "B5", "SUM(B3:B3)", "30");

    for (n, a) in [
        ("Cfg_BaseYear", "Settings!$B$1"),
        ("Cfg_NewInvRate", "Settings!$B$2"),
        ("Exp_Item", "Expenditures!$A$2:$A$3"),
        ("Exp_Element", "Expenditures!$B$2:$B$3"),
        ("Exp_GLName", "Expenditures!$C$2:$C$3"),
        ("Exp_Reason", "Expenditures!$D$2:$D$3"),
        ("Exp_RFS", "Expenditures!$E$2:$G$3"),
        ("Exp_RFSYears", "Expenditures!$E$1:$G$1"),
        ("Exp_Revised", "Expenditures!$H$2:$J$3"),
        ("Exp_RevisedYears", "Expenditures!$H$1:$J$1"),
        ("MTC_Title", "Medium_Term_Cashflow!$A$1"),
        ("MTC_Years", "Medium_Term_Cashflow!$B$3:$D$3"),
        ("MTC_CashSection", "Medium_Term_Cashflow!$A$3:$D$6"),
        ("MTC_ComparisonSection", "Medium_Term_Cashflow!$A$8:$D$9"),
        ("Inv_Title", "Schedule_of_Investments!$A$1"),
        ("Inv_Totals", "Schedule_of_Investments!$B$5:$D$5"),
    ] {
        name(&mut wb, n, a);
    }
    wb
}

fn run(test: &str) -> Workbook {
    let d = std::env::temp_dir().join(format!("plan-tests-{}-{test}", std::process::id()));
    let _ = std::fs::remove_dir_all(&d);
    std::fs::create_dir_all(&d).unwrap();
    let src: PathBuf = d.join("Spatac-20261121.xlsx");
    let out = d.join("InvestmentPlan-20261121.xlsx");
    umya_spreadsheet::writer::xlsx::write(&spatac_shaped(), &src).unwrap();
    extract(&src, &out, "2026-11-21").unwrap();
    umya_spreadsheet::reader::xlsx::read(&out).unwrap()
}

#[test]
fn tabs_and_order() {
    let wb = run("tabs");
    let names: Vec<&str> = wb.sheet_collection().iter().map(|s| s.name()).collect();
    assert_eq!(
        names,
        ["Cover", "Medium_Term_Cashflow", "Schedule_of_Investments"]
    );
    let cover = wb.sheet_by_name("Cover").unwrap();
    let text: Vec<String> = cover
        .cells()
        .iter()
        .map(|c| c.value().into_owned())
        .collect();
    assert!(text.iter().any(|t| t == "Spatac-20261121.xlsx"), "{text:?}");
    assert!(text.iter().any(|t| t == "2026-11-21"));
    assert!(text.iter().any(|t| t == "2026"));
    assert!(
        text.iter().any(|t| t.contains(
            "reasons for those changes are described in the Spatac workbooks for the current and \
             previous years"
        )),
        "{text:?}"
    );
}

#[test]
fn references_to_other_tabs_become_values() {
    let wb = run("values");
    let mtc = wb.sheet_by_name("Medium_Term_Cashflow").unwrap();
    let b3 = mtc.cell("B3").unwrap();
    assert!(!b3.is_formula(), "Cfg_BaseYear is on Settings");
    assert_eq!(b3.value(), "2026");
    assert_eq!(mtc.cell("C3").unwrap().formula(), "B3+1");
    assert_eq!(mtc.cell("C5").unwrap().formula(), "B5*2");
    let b6 = mtc.cell("B6").unwrap();
    assert!(!b6.is_formula(), "Expenditures is not in the plan");
    assert_eq!(b6.value(), "-50");
    // A reference to the other kept tab stays a formula.
    assert_eq!(
        mtc.cell("B8").unwrap().formula(),
        "SUM(Schedule_of_Investments!D3:D3)"
    );
    let inv = wb.sheet_by_name("Schedule_of_Investments").unwrap();
    assert!(!inv.cell("B3").unwrap().is_formula());
    assert_eq!(inv.cell("B5").unwrap().formula(), "SUM(B3:B3)");
}

#[test]
fn rows_outside_the_plan_are_dropped() {
    let wb = run("trim");
    let mtc = wb.sheet_by_name("Medium_Term_Cashflow").unwrap();
    assert!(mtc.cell("A20").is_none_or(|c| c.value().is_empty()));
    assert!(mtc.cell("B20").is_none_or(|c| c.value().is_empty()));
}

#[test]
fn everything_is_locked_and_protected() {
    let wb = run("locked");
    for ws in wb.sheet_collection() {
        assert!(
            ws.sheet_protection().is_some_and(|p| p.sheet()),
            "{} not protected",
            ws.name()
        );
        for c in ws.cells() {
            assert!(
                c.style().protection().is_none_or(|p| p.locked()),
                "{}!{} unlocked",
                ws.name(),
                c.coordinate()
            );
        }
    }
}
