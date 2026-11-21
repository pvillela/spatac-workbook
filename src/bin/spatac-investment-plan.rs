//! Extracts the Investment Plan from a Spatac workbook into a separate workbook.

use clap::Parser;
use spatac_workbook::investment_plan::{self, default_output};
use spatac_workbook::notices::THIRD_PARTY_NOTICES;
use std::path::PathBuf;
use std::process::ExitCode;

/// Extracts the Investment Plan from a Spatac workbook into a separate workbook.
///
/// The output holds a cover sheet, the Medium_Term_Cashflow and Schedule_of_Investments tabs,
/// and a list of changes from the Reserve Fund Study. Formulas that refer to other tabs become
/// their values; every cell is locked. Save the source workbook (so its values are up to date)
/// before running this.
#[derive(Parser)]
#[command(version, about, long_about)]
struct Args {
    /// The Spatac workbook, e.g. Spatac-20261121.xlsx.
    #[arg(required_unless_present = "third_party_notices")]
    source: Option<PathBuf>,
    /// Where to write the plan. Defaults to InvestmentPlan-YYYYMMDD.xlsx beside the source,
    /// with the date taken from the source's name.
    #[arg(short, long)]
    output: Option<PathBuf>,
    /// Replace the output file if it exists.
    #[arg(long)]
    force: bool,
    /// Print the third-party notices and exit.
    #[arg(long)]
    third_party_notices: bool,
}

fn main() -> ExitCode {
    let args = Args::parse();
    if args.third_party_notices {
        print!("{THIRD_PARTY_NOTICES}");
        return ExitCode::SUCCESS;
    }
    let Some(source) = args.source else {
        unreachable!("clap requires the source");
    };
    let output = match args.output.or_else(|| default_output(&source)) {
        Some(o) => o,
        None => {
            eprintln!(
                "spatac-investment-plan: {} has no YYYYMMDD date in its name; give --output",
                source.display()
            );
            return ExitCode::FAILURE;
        }
    };
    if output.exists() && !args.force {
        eprintln!(
            "spatac-investment-plan: {} exists; pass --force to replace it",
            output.display()
        );
        return ExitCode::FAILURE;
    }
    match investment_plan::extract(&source, &output, &investment_plan::today()) {
        Ok(s) => {
            println!("Wrote {}", s.output.display());
            println!(
                "  {} formulas kept, {} replaced by their values",
                s.formulas_kept, s.formulas_to_values
            );
            println!(
                "  {} changes from the RFS, {} without a reason",
                s.changes, s.changes_without_reason
            );
            ExitCode::SUCCESS
        }
        Err(e) => {
            eprintln!("spatac-investment-plan: {e}");
            ExitCode::FAILURE
        }
    }
}
