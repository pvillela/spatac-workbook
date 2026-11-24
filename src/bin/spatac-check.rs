//! Checks the integrity of a Spatac workbook.

use clap::Parser;
use spatac_workbook::check::check;
use spatac_workbook::notices::THIRD_PARTY_NOTICES;
use std::path::PathBuf;
use std::process::ExitCode;

/// Checks the integrity of a Spatac workbook and lists every problem found.
///
/// It reads the saved file without recalculating it: missing formulas (for example in a row
/// inserted into an Adjustments block), broken or missing names, cells showing an error,
/// Adjustments inputs without their validation, and protection that does not match a cell's
/// role. Save the workbook in Excel or Calc before running it. The exit status is 1 when a
/// problem is found.
#[derive(Parser)]
#[command(version, about, long_about)]
struct Args {
    /// The Spatac workbook, e.g. Spatac-20261124.xlsx.
    #[arg(required_unless_present = "third_party_notices")]
    workbook: Option<PathBuf>,
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
    let Some(path) = args.workbook else {
        unreachable!("clap requires the workbook");
    };
    match check(&path) {
        Ok(problems) if problems.is_empty() => {
            println!("{}: no problems found", path.display());
            ExitCode::SUCCESS
        }
        Ok(problems) => {
            println!("{}: {} problem(s)", path.display(), problems.len());
            for p in &problems {
                println!("- {p}");
            }
            ExitCode::FAILURE
        }
        Err(e) => {
            eprintln!("spatac-check: {e}");
            ExitCode::from(2)
        }
    }
}
