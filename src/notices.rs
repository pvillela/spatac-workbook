//! Third-party notices embedded by `build.rs`.

/// The notices for the crates linked into this binary. A release build carries the generated
/// `THIRD-PARTY-NOTICES.md`; any other build carries a placeholder saying how to produce it.
pub const THIRD_PARTY_NOTICES: &str =
    include_str!(concat!(env!("OUT_DIR"), "/third-party-notices.md"));
