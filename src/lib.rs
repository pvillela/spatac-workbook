//! Support code for the Spatac workbook command-line tools.
//!
//! - [`diff`] compares two workbooks and reports the differences as Markdown. It knows nothing
//!   about the Spatac workbook and works on any `.xlsx` file.
//! - [`investment_plan`] extracts the Investment Plan from a Spatac workbook.
//! - [`check`] checks the integrity of a Spatac workbook.

pub mod check;
pub mod diff;
pub mod investment_plan;
pub mod notices;
