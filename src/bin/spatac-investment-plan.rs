//! Extracts the Investment Plan from a Spatac workbook into a separate workbook.

use clap::Parser;
use spatac_workbook::notices::THIRD_PARTY_NOTICES;
use std::process::ExitCode;

#[derive(Parser)]
#[command(version, about)]
struct Args {
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
    eprintln!("spatac-investment-plan: not implemented yet");
    ExitCode::FAILURE
}
