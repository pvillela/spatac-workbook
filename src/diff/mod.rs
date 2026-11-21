//! Workbook comparison.
//!
//! [`diff_files`] reads two `.xlsx` workbooks and reports, as a [`Report`], what differs:
//!
//! - tabs added, deleted, renamed (recognised by content) and reordered;
//! - defined names;
//! - for each tab present in both: inserted, deleted and moved rows and columns, then every cell
//!   whose formula, constant or comment changed. Formulas are compared after their references
//!   have been rewritten through the row and column alignment, so a formula that only moved, or
//!   whose references shifted with an inserted row, is not reported. Cells changed the same way
//!   down a column are reported as one entry;
//! - merged cells, conditional formats, data validation, sheet protection, unlocked cells,
//!   hidden rows, columns and tabs, and frozen panes;
//! - with [`Options::values`], recalculated results of formulas that did not change;
//! - with [`Options::formats`], cell formats, column widths and row heights.
//!
//! [`markdown::render`] writes the report as Markdown. Nothing here knows about the Spatac
//! workbook.

pub mod align;
pub mod compare;
pub mod formula;
pub mod markdown;
pub mod model;

pub use compare::{Options, Report};

use std::path::Path;

/// Compares two workbook files.
pub fn diff_files(old: &Path, new: &Path, options: Options) -> Result<Report, String> {
    let a = model::load(old, options.formats)?;
    let b = model::load(new, options.formats)?;
    Ok(compare::compare(&a, &b, options))
}
