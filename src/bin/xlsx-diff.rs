//! Compares two `.xlsx` workbooks and writes the differences as Markdown.

use clap::Parser;
use spatac_workbook::diff::{self, Options, markdown};
use spatac_workbook::notices::THIRD_PARTY_NOTICES;
use std::path::PathBuf;
use std::process::ExitCode;

/// Compares two versions of a workbook and writes the differences as Markdown.
///
/// Reports tabs added, deleted, renamed and reordered; defined names; and, for each tab in both
/// versions, inserted, deleted and moved rows and columns, changed formulas, constants and
/// comments, and structural changes (merged cells, conditional formats, data validation,
/// protection, unlocked cells, hidden rows and columns, frozen panes), and changed cell formats,
/// column widths and row heights.
#[derive(Parser)]
#[command(version, about, long_about)]
struct Args {
    /// The older workbook.
    #[arg(required_unless_present = "third_party_notices")]
    old: Option<PathBuf>,
    /// The newer workbook.
    #[arg(required_unless_present = "third_party_notices")]
    new: Option<PathBuf>,
    /// Write the report to this file instead of standard output.
    #[arg(short, long)]
    output: Option<PathBuf>,
    /// Also report recalculated results of formulas that did not change.
    #[arg(long)]
    values: bool,
    /// Leave out changed cell formats, column widths and row heights.
    #[arg(long)]
    no_formats: bool,
    /// Formats are reported by default; kept so that older command lines still work.
    #[arg(long, hide = true, conflicts_with = "no_formats")]
    formats: bool,
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
    let (Some(old), Some(new)) = (args.old, args.new) else {
        unreachable!("clap requires both paths");
    };
    let options = Options {
        values: args.values,
        formats: args.formats || !args.no_formats,
    };
    let report = match diff::diff_files(&old, &new, options) {
        Ok(r) => r,
        Err(e) => {
            eprintln!("xlsx-diff: {e}");
            return ExitCode::FAILURE;
        }
    };
    let label = |p: &PathBuf| {
        p.file_name().map_or_else(
            || p.display().to_string(),
            |f| f.to_string_lossy().into_owned(),
        )
    };
    let md = markdown::render(&report, &label(&old), &label(&new));
    match args.output {
        Some(path) => {
            if let Err(e) = std::fs::write(&path, md) {
                eprintln!("xlsx-diff: cannot write {}: {e}", path.display());
                return ExitCode::FAILURE;
            }
        }
        None => print!("{md}"),
    }
    ExitCode::SUCCESS
}
