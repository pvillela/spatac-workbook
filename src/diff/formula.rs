//! Cell references inside formula text.
//!
//! A formula is split into literal text and references. That is enough to do the two things the
//! diff needs: render a formula in relative (R1C1-style) form, so that formulas filled down a
//! column compare equal, and rewrite a formula's references through a row and column mapping, so
//! that a formula whose rows merely moved is not reported as changed.
//!
//! This is deliberately not a full formula parser. References are recognised by shape, outside
//! string literals, with checks on the characters either side so that function names such as
//! `LOG10(` and defined names such as `Exp_RFS` are left alone.

use regex::Regex;
use std::sync::LazyLock;

/// One end of a reference. A whole-column reference has no row and a whole-row reference has no
/// column.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct RefPart {
    pub col: Option<u32>,
    pub col_abs: bool,
    pub row: Option<u32>,
    pub row_abs: bool,
}

/// A cell, range, whole-column or whole-row reference, optionally qualified by a sheet name.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Reference {
    /// The sheet name, unquoted. `None` means the sheet holding the formula.
    pub sheet: Option<String>,
    pub start: RefPart,
    pub end: Option<RefPart>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Token {
    Text(String),
    Ref(Reference),
}

static REF_RE: LazyLock<Regex> = LazyLock::new(|| {
    Regex::new(
        r"(?x)
        (?: (?P<sheet> '(?:[^']|'')+' | [A-Za-z_\\][A-Za-z0-9_.]* ) ! )?
        (?P<body>
            \$?[A-Za-z]{1,3}\$?[0-9]+ (?: : \$?[A-Za-z]{1,3}\$?[0-9]+ )?
          | \$?[A-Za-z]{1,3} : \$?[A-Za-z]{1,3}
          | \$?[0-9]+ : \$?[0-9]+
        )",
    )
    .expect("the reference pattern is valid")
});

/// Splits formula text (without the leading `=`) into literal text and references.
pub fn tokenize(formula: &str) -> Vec<Token> {
    let mut tokens = Vec::new();
    let mut text = String::new();
    let bytes = formula.as_bytes();
    let mut i = 0;
    while i < formula.len() {
        if bytes[i] == b'"' {
            // A string literal runs to the next unpaired quote.
            let start = i;
            i += 1;
            while i < formula.len() {
                if bytes[i] == b'"' {
                    if i + 1 < formula.len() && bytes[i + 1] == b'"' {
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                i += 1;
            }
            text.push_str(&formula[start..i]);
            continue;
        }
        let next_quote = formula[i..].find('"').map_or(formula.len(), |q| i + q);
        let segment_end = next_quote;
        let mut pos = i;
        while pos < segment_end {
            let Some(caps) = REF_RE.captures_at(&formula[..segment_end], pos) else {
                break;
            };
            let whole = caps.get(0).expect("group 0 always matches");
            if !boundary_ok(formula, whole.start(), whole.end()) {
                // Not a reference here; it may still start one character later.
                let advance = formula[whole.start()..]
                    .chars()
                    .next()
                    .map_or(1, char::len_utf8);
                text.push_str(&formula[pos..whole.start() + advance]);
                pos = whole.start() + advance;
                continue;
            }
            text.push_str(&formula[pos..whole.start()]);
            match parse_reference(caps.name("sheet").map(|m| m.as_str()), &caps["body"]) {
                Some(reference) => {
                    if !text.is_empty() {
                        tokens.push(Token::Text(std::mem::take(&mut text)));
                    }
                    tokens.push(Token::Ref(reference));
                }
                None => text.push_str(whole.as_str()),
            }
            pos = whole.end();
        }
        text.push_str(&formula[pos..segment_end]);
        i = segment_end;
    }
    if !text.is_empty() {
        tokens.push(Token::Text(text));
    }
    tokens
}

/// A reference must not be glued to an identifier, a number or a function call.
fn boundary_ok(formula: &str, start: usize, end: usize) -> bool {
    let before = formula[..start].chars().next_back();
    let after = formula[end..].chars().next();
    let ident = |c: char| c.is_ascii_alphanumeric() || c == '_' || c == '.' || c == '\\';
    if before.is_some_and(|c| ident(c) || c == '$' || c == '\'' || c == '!') {
        return false;
    }
    if after.is_some_and(|c| ident(c) || c == '(' || c == '!' || c == '$' || c == ':') {
        return false;
    }
    true
}

fn parse_reference(sheet: Option<&str>, body: &str) -> Option<Reference> {
    let sheet = sheet.map(|s| {
        if let Some(inner) = s.strip_prefix('\'').and_then(|s| s.strip_suffix('\'')) {
            inner.replace("''", "'")
        } else {
            s.to_owned()
        }
    });
    let (a, b) = match body.split_once(':') {
        Some((a, b)) => (a, Some(b)),
        None => (body, None),
    };
    let start = parse_part(a)?;
    let end = match b {
        Some(b) => Some(parse_part(b)?),
        None => None,
    };
    Some(Reference { sheet, start, end })
}

fn parse_part(s: &str) -> Option<RefPart> {
    let mut chars = s.chars().peekable();
    let mut part = RefPart {
        col: None,
        col_abs: false,
        row: None,
        row_abs: false,
    };
    let mut abs = false;
    if chars.peek() == Some(&'$') {
        abs = true;
        chars.next();
    }
    let mut letters = String::new();
    while let Some(&c) = chars.peek() {
        if c.is_ascii_alphabetic() {
            letters.push(c.to_ascii_uppercase());
            chars.next();
        } else {
            break;
        }
    }
    if !letters.is_empty() {
        part.col = Some(column_number(&letters)?);
        part.col_abs = abs;
        abs = false;
        if chars.peek() == Some(&'$') {
            abs = true;
            chars.next();
        }
    }
    let digits: String = chars.collect();
    if !digits.is_empty() {
        let row: u32 = digits.parse().ok()?;
        if row == 0 || row > 1_048_576 {
            return None;
        }
        part.row = Some(row);
        part.row_abs = abs;
    } else if abs {
        return None;
    }
    if part.col.is_none() && part.row.is_none() {
        return None;
    }
    Some(part)
}

/// `A` is 1, `Z` is 26, `AA` is 27. Columns beyond `XFD` are not columns.
pub fn column_number(letters: &str) -> Option<u32> {
    let mut n: u32 = 0;
    for c in letters.chars() {
        if !c.is_ascii_alphabetic() {
            return None;
        }
        n = n * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    (1..=16_384).contains(&n).then_some(n)
}

/// 1 is `A`, 27 is `AA`.
pub fn column_letters(mut n: u32) -> String {
    let mut s = Vec::new();
    while n > 0 {
        let r = (n - 1) % 26;
        s.push(b'A' + r as u8);
        n = (n - 1) / 26;
    }
    s.reverse();
    String::from_utf8(s).expect("ASCII letters")
}

/// `B5` for row 5, column 2.
pub fn cell_name(row: u32, col: u32) -> String {
    format!("{}{}", column_letters(col), row)
}

/// Renders a formula with each relative reference written as an offset from `(row, col)`, so
/// that the same formula filled across a block renders identically in every cell.
pub fn relative_form(tokens: &[Token], row: u32, col: u32) -> String {
    let mut out = String::new();
    for token in tokens {
        match token {
            Token::Text(t) => out.push_str(&strip_spaces(t)),
            Token::Ref(r) => {
                if let Some(sheet) = &r.sheet {
                    out.push_str(&quote_sheet(sheet));
                    out.push('!');
                }
                out.push_str(&relative_part(&r.start, row, col));
                if let Some(end) = &r.end {
                    out.push(':');
                    out.push_str(&relative_part(end, row, col));
                }
            }
        }
    }
    out
}

fn relative_part(p: &RefPart, row: u32, col: u32) -> String {
    let mut s = String::new();
    if let Some(r) = p.row {
        if p.row_abs {
            s.push_str(&format!("R{r}"));
        } else {
            s.push_str(&format!("R[{}]", i64::from(r) - i64::from(row)));
        }
    }
    if let Some(c) = p.col {
        if p.col_abs {
            s.push_str(&format!("C{c}"));
        } else {
            s.push_str(&format!("C[{}]", i64::from(c) - i64::from(col)));
        }
    }
    s
}

/// Where a row or column of the old workbook is in the new one.
pub trait Remap {
    /// The new index of an old row or column, if it survived.
    fn exact(&self, old: u32) -> Option<u32>;
    /// For the start of a range: the first surviving index at or after `old`.
    fn forward(&self, old: u32) -> Option<u32>;
    /// For the end of a range: the last surviving index at or before `old`.
    fn backward(&self, old: u32) -> Option<u32>;
}

/// The row and column mappings of one sheet, and its new name.
pub struct SheetMap<'a> {
    pub new_name: &'a str,
    pub rows: &'a dyn Remap,
    pub cols: &'a dyn Remap,
}

/// Rewrites a formula of the old workbook into the coordinates of the new one.
///
/// `lookup` resolves a sheet name of the old workbook, or `None` for the formula's own sheet. A
/// sheet it cannot resolve (one deleted in the new workbook) leaves its references untouched. A
/// single-cell reference to a deleted row or column becomes `#REF!`, as it would in the
/// spreadsheet application, and so does a range all of whose rows or columns were deleted.
pub fn remap<'a>(
    tokens: &[Token],
    lookup: &dyn Fn(Option<&str>) -> Option<SheetMap<'a>>,
) -> String {
    let mut out = String::new();
    for token in tokens {
        match token {
            Token::Text(t) => out.push_str(&strip_spaces(t)),
            Token::Ref(r) => {
                let Some(map) = lookup(r.sheet.as_deref()) else {
                    out.push_str(&render_reference(r, r.sheet.as_deref()));
                    continue;
                };
                let sheet = r.sheet.as_ref().map(|_| map.new_name);
                let mapped = match &r.end {
                    None => map_part(&r.start, &map, Direction::Exact).map(|start| Reference {
                        sheet: None,
                        start,
                        end: None,
                    }),
                    Some(end) => {
                        let start = map_part(&r.start, &map, Direction::Forward);
                        let end = map_part(end, &map, Direction::Backward);
                        match (start, end) {
                            // Both ends in deleted rows or columns: the range is gone, and the
                            // ends cross (S7:S250 with S deleted would read U7:R250).
                            (Some(start), Some(end))
                                if start.col > end.col || start.row > end.row =>
                            {
                                None
                            }
                            (Some(start), Some(end)) => Some(Reference {
                                sheet: None,
                                start,
                                end: Some(end),
                            }),
                            _ => None,
                        }
                    }
                };
                match mapped {
                    Some(m) => out.push_str(&render_reference(&m, sheet)),
                    None => {
                        if let Some(s) = sheet {
                            out.push_str(&quote_sheet(s));
                            out.push('!');
                        }
                        out.push_str("#REF!");
                    }
                }
            }
        }
    }
    out
}

#[derive(Clone, Copy)]
enum Direction {
    Exact,
    Forward,
    Backward,
}

fn map_part(p: &RefPart, map: &SheetMap, dir: Direction) -> Option<RefPart> {
    let pick = |remap: &dyn Remap, old: u32| match dir {
        Direction::Exact => remap.exact(old),
        Direction::Forward => remap.forward(old),
        Direction::Backward => remap.backward(old),
    };
    Some(RefPart {
        col: match p.col {
            Some(c) => Some(pick(map.cols, c)?),
            None => None,
        },
        col_abs: p.col_abs,
        row: match p.row {
            Some(r) => Some(pick(map.rows, r)?),
            None => None,
        },
        row_abs: p.row_abs,
    })
}

/// Writes formula text back out, with insignificant spaces removed, so that two formulas compare
/// equal whenever they say the same thing.
pub fn normalized(tokens: &[Token]) -> String {
    let mut out = String::new();
    for token in tokens {
        match token {
            Token::Text(t) => out.push_str(&strip_spaces(t)),
            Token::Ref(r) => out.push_str(&render_reference(r, r.sheet.as_deref())),
        }
    }
    out
}

fn render_reference(r: &Reference, sheet: Option<&str>) -> String {
    let mut s = String::new();
    if let Some(sheet) = sheet {
        s.push_str(&quote_sheet(sheet));
        s.push('!');
    }
    s.push_str(&render_part(&r.start));
    if let Some(end) = &r.end {
        s.push(':');
        s.push_str(&render_part(end));
    }
    s
}

fn render_part(p: &RefPart) -> String {
    let mut s = String::new();
    if let Some(c) = p.col {
        if p.col_abs {
            s.push('$');
        }
        s.push_str(&column_letters(c));
    }
    if let Some(r) = p.row {
        if p.row_abs {
            s.push('$');
        }
        s.push_str(&r.to_string());
    }
    s
}

/// Quotes a sheet name when a formula would need it quoted.
pub fn quote_sheet(name: &str) -> String {
    let plain = !name.is_empty()
        && name
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || c == '_' || c == '.')
        && !name.starts_with(|c: char| c.is_ascii_digit());
    if plain {
        name.to_owned()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

/// Removes spaces outside string literals. Literal text tokens never contain a whole string
/// literal split across tokens, because [`tokenize`] keeps each literal inside one text token.
fn strip_spaces(t: &str) -> String {
    let mut out = String::with_capacity(t.len());
    let mut in_string = false;
    for c in t.chars() {
        if c == '"' {
            in_string = !in_string;
        }
        if in_string || !c.is_whitespace() {
            out.push(c);
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn refs(f: &str) -> Vec<String> {
        tokenize(f)
            .into_iter()
            .filter_map(|t| match t {
                Token::Ref(r) => Some(render_reference(&r, r.sheet.as_deref())),
                Token::Text(_) => None,
            })
            .collect()
    }

    #[test]
    fn recognises_references_and_leaves_names_alone() {
        assert_eq!(
            refs("SUM(B5:B12) + B22*LOG10(Exp_RFS)+'fvalues spatac'!$K$7+A:A+3:4"),
            vec!["B5:B12", "B22", "'fvalues spatac'!$K$7", "A:A", "3:4"]
        );
        assert_eq!(refs("1E5+X1Y"), Vec::<String>::new());
        assert_eq!(refs("\"A1\"&B2"), vec!["B2"]);
        assert_eq!(
            refs("cashflow_spatac!M17:M22"),
            vec!["cashflow_spatac!M17:M22"]
        );
    }

    #[test]
    fn relative_form_matches_across_a_fill() {
        let a = relative_form(&tokenize("B5*$C$1"), 5, 4);
        let b = relative_form(&tokenize("B6*$C$1"), 6, 4);
        assert_eq!(a, b);
        assert_eq!(a, "R[0]C[-2]*R1C3");
    }

    struct Shift(u32);
    impl Remap for Shift {
        fn exact(&self, old: u32) -> Option<u32> {
            Some(if old >= self.0 { old + 1 } else { old })
        }
        fn forward(&self, old: u32) -> Option<u32> {
            self.exact(old)
        }
        fn backward(&self, old: u32) -> Option<u32> {
            self.exact(old)
        }
    }
    struct Same;
    impl Remap for Same {
        fn exact(&self, old: u32) -> Option<u32> {
            Some(old)
        }
        fn forward(&self, old: u32) -> Option<u32> {
            Some(old)
        }
        fn backward(&self, old: u32) -> Option<u32> {
            Some(old)
        }
    }

    #[test]
    fn remap_follows_an_inserted_row() {
        let rows = Shift(5);
        let lookup = |s: Option<&str>| {
            s.is_none().then_some(SheetMap {
                new_name: "S",
                rows: &rows as &dyn Remap,
                cols: &Same as &dyn Remap,
            })
        };
        assert_eq!(
            remap(&tokenize("SUM(A1:A10) + Other!A7"), &lookup),
            "SUM(A1:A11)+Other!A7"
        );
    }
}
