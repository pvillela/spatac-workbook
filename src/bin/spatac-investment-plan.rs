//! Extracts the Investment Plan from a Spatac workbook into a separate workbook.

use clap::Parser;
use spatac_workbook::investment_plan::{self, default_output};
use spatac_workbook::notices::THIRD_PARTY_NOTICES;
use std::path::PathBuf;
use std::process::ExitCode;

/// Extracts the Investment Plan from a Spatac workbook into a separate workbook.
///
/// The output holds a cover sheet and the Medium_Term_Cashflow and Schedule_of_Investments tabs.
/// Formulas that refer to other tabs become their values; every cell is locked. Save the source
/// workbook (so its values are up to date) before running this.
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
    /// The date to show as "Generated" on the cover, as YYYY-MM-DD. Defaults to today, in this
    /// computer's time zone.
    #[arg(long, value_parser = investment_plan::parse_date)]
    date: Option<String>,
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
    let date = args.date.unwrap_or_else(investment_plan::today);
    match investment_plan::extract(&source, &output, &date) {
        Ok(s) => {
            println!("Wrote {}", s.output.display());
            println!(
                "  {} formulas kept, {} replaced by their values",
                s.formulas_kept, s.formulas_to_values
            );
            ExitCode::SUCCESS
        }
        Err(e) => {
            eprintln!("spatac-investment-plan: {e}");
            ExitCode::FAILURE
        }
    }
}
