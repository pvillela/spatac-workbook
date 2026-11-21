//! Renders a [`Report`] as Markdown.

use super::compare::{CellChange, NameChange, Report};
use std::fmt::Write;

/// Renders the report. `old` and `new` label the two versions, usually their file names.
pub fn render(report: &Report, old: &str, new: &str) -> String {
    let mut md = String::new();
    let w = &mut md;
    let _ = writeln!(w, "# Workbook differences\n");
    let _ = writeln!(w, "- Old: `{}`", old);
    let _ = writeln!(w, "- New: `{}`\n", new);

    if report.is_empty() {
        let _ = writeln!(w, "No differences.");
        return md;
    }

    let _ = writeln!(w, "## Summary\n");
    let list = |names: &[String]| {
        names
            .iter()
            .map(|n| format!("`{n}`"))
            .collect::<Vec<_>>()
            .join(", ")
    };
    if !report.added_sheets.is_empty() {
        let _ = writeln!(w, "- Tabs added: {}", list(&report.added_sheets));
    }
    if !report.deleted_sheets.is_empty() {
        let _ = writeln!(w, "- Tabs deleted: {}", list(&report.deleted_sheets));
    }
    for (a, b) in &report.renamed_sheets {
        let _ = writeln!(w, "- Tab renamed: `{a}` → `{b}`");
    }
    if let Some((before, after)) = &report.order {
        let _ = writeln!(w, "- Tab order changed:");
        for (name, a, b) in &report.reordered_sheets {
            let _ = writeln!(
                w,
                "  - `{name}` moved from position {a} to {b} among the tabs in both versions"
            );
        }
        let _ = writeln!(w, "  - Old order: {}", list(before));
        let _ = writeln!(w, "  - New order: {}", list(after));
    }
    if !report.sheets.is_empty() {
        let changed: Vec<String> = report.sheets.iter().map(|s| s.name.clone()).collect();
        let _ = writeln!(w, "- Tabs changed: {}", list(&changed));
    }
    if !report.unchanged_sheets.is_empty() {
        let _ = writeln!(w, "- Tabs unchanged: {}", list(&report.unchanged_sheets));
    }
    if !report.names.is_empty() {
        let _ = writeln!(w, "- Defined names changed: {}", report.names.len());
    }
    let _ = writeln!(w);

    if !report.names.is_empty() {
        let _ = writeln!(w, "## Defined names\n");
        for change in &report.names {
            let _ = match change {
                NameChange::Added { name, refers_to } => {
                    writeln!(w, "- Added `{name}` = {}", code(refers_to))
                }
                NameChange::Removed { name, refers_to } => {
                    writeln!(w, "- Removed `{name}` (was {})", code(refers_to))
                }
                NameChange::Changed {
                    name,
                    before,
                    after,
                } => writeln!(w, "- Changed `{name}`: {} → {}", code(before), code(after)),
            };
        }
        let _ = writeln!(w);
    }

    for sheet in &report.sheets {
        match &sheet.renamed_from {
            Some(from) => {
                let _ = writeln!(w, "## Tab `{}` (renamed from `{from}`)\n", sheet.name);
            }
            None => {
                let _ = writeln!(w, "## Tab `{}`\n", sheet.name);
            }
        }
        if !sheet.structure.is_empty() {
            let _ = writeln!(w, "### Structure\n");
            for s in &sheet.structure {
                let _ = writeln!(w, "- {s}");
            }
            let _ = writeln!(w);
        }
        if !sheet.cells.is_empty() {
            let cells: u32 = sheet.cells.iter().map(|c| c.count).sum();
            let _ = writeln!(w, "### Cell changes\n");
            let _ = writeln!(
                w,
                "<details open>\n<summary>{} {} in {} {}</summary>\n",
                cells,
                plural(cells, "cell", "cells"),
                sheet.cells.len(),
                plural(sheet.cells.len() as u32, "entry", "entries")
            );
            table(w, &sheet.cells);
            let _ = writeln!(w, "\n</details>\n");
        }
        if !sheet.values.is_empty() {
            let cells: u32 = sheet.values.iter().map(|c| c.count).sum();
            let _ = writeln!(w, "### Recalculated values\n");
            let _ = writeln!(
                w,
                "<details>\n<summary>{} {} whose formula did not change</summary>\n",
                cells,
                plural(cells, "cell", "cells")
            );
            table(w, &sheet.values);
            let _ = writeln!(w, "\n</details>\n");
        }
        if !sheet.formats.is_empty() {
            let _ = writeln!(w, "### Formats\n");
            for s in &sheet.formats {
                let _ = writeln!(w, "- {s}");
            }
            let _ = writeln!(w);
        }
    }
    md
}

fn plural(n: u32, one: &'static str, many: &'static str) -> &'static str {
    if n == 1 { one } else { many }
}

fn table(w: &mut String, changes: &[CellChange]) {
    let _ = writeln!(w, "| New | Old | Change | Before | After |");
    let _ = writeln!(w, "|---|---|---|---|---|");
    for c in changes {
        let new = c.new.clone().unwrap_or_else(|| "—".to_owned());
        let old = match (&c.old, &c.new) {
            (Some(o), Some(n)) if o == n => String::new(),
            (Some(o), _) => o.clone(),
            (None, _) => "—".to_owned(),
        };
        let mut kind = c.kind.label().to_owned();
        if c.count > 1 {
            kind.push_str(&format!(" ×{}", c.count));
        }
        let _ = writeln!(
            w,
            "| {} | {} | {} | {} | {} |",
            new,
            old,
            kind,
            cell(&c.before),
            cell(&c.after)
        );
    }
}

/// A table cell: code for formulas and values, plain text for comments, never breaking the row.
fn cell(s: &str) -> String {
    if s.is_empty() {
        return String::new();
    }
    code(s)
}

fn code(s: &str) -> String {
    let flat = s
        .replace("\r\n", "\n")
        .replace('\n', " ⏎ ")
        .replace('|', "\\|");
    if flat.contains('`') {
        format!("`` {flat} ``")
    } else {
        format!("`{flat}`")
    }
}
