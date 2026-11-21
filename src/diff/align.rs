//! Aligns the rows (or columns) of two versions of a sheet.
//!
//! Exact matches come from a Myers diff of per-row signatures. A row that was edited no longer
//! matches exactly, so inside each block the Myers diff reports as replaced (deletions and
//! insertions in the same gap between matches), rows are paired by similarity instead: an edited row is the same row, changed, rather than one row deleted and
//! another inserted. What remains unpaired is inserted or deleted, except that an unpaired old row
//! whose content reappears exactly as an unpaired new row is reported as moved.

use super::formula::Remap;
use similar::{Algorithm, DiffOp, capture_diff_slices};
use std::collections::{BTreeMap, HashMap};

/// Rows (or columns) whose similarity is below this are not paired.
const PAIR_THRESHOLD: f64 = 0.5;

#[derive(Debug, Default)]
pub struct Alignment {
    /// Old index to new index, for every old index that survives (paired or moved).
    pub old_to_new: BTreeMap<u32, u32>,
    pub new_to_old: BTreeMap<u32, u32>,
    /// Old indices with no counterpart.
    pub deleted: Vec<u32>,
    /// New indices with no counterpart.
    pub inserted: Vec<u32>,
    /// (old, new) pairs that moved, as opposed to merely shifting with inserts and deletes.
    pub moved: Vec<(u32, u32)>,
}

impl Alignment {
    /// Indices 1..=n mapped onto themselves.
    pub fn identity(n: u32) -> Self {
        let mut a = Alignment::default();
        for i in 1..=n {
            a.old_to_new.insert(i, i);
            a.new_to_old.insert(i, i);
        }
        a
    }

    /// The aligned (old, new) pairs in new order, followed by nothing else.
    pub fn pairs(&self) -> impl Iterator<Item = (u32, u32)> + '_ {
        self.new_to_old.iter().map(|(&n, &o)| (o, n))
    }

    /// Lies beyond the compared area: such an index maps to itself shifted by the net count of
    /// rows inserted and deleted above it.
    fn beyond(&self, old: u32) -> Option<u32> {
        let (&last_old, &last_new) = self.old_to_new.iter().next_back()?;
        (old > last_old).then(|| old - last_old + last_new)
    }
}

impl Remap for Alignment {
    fn exact(&self, old: u32) -> Option<u32> {
        self.old_to_new
            .get(&old)
            .copied()
            .or_else(|| self.beyond(old))
            .or_else(|| self.old_to_new.is_empty().then_some(old))
    }

    fn forward(&self, old: u32) -> Option<u32> {
        self.old_to_new
            .range(old..)
            .next()
            .map(|(_, &n)| n)
            .or_else(|| self.beyond(old))
            .or_else(|| self.old_to_new.is_empty().then_some(old))
    }

    fn backward(&self, old: u32) -> Option<u32> {
        if let Some(n) = self.beyond(old) {
            return Some(n);
        }
        self.old_to_new
            .range(..=old)
            .next_back()
            .map(|(_, &n)| n)
            .or_else(|| self.old_to_new.is_empty().then_some(old))
    }
}

/// Aligns sequences `old` and `new` (index 0 is row or column 1), given exact signatures and a
/// similarity function for pairing edited entries. Empty signatures are blank rows: they align
/// with each other but are never reported as moved.
pub fn align(
    old: &[u64],
    new: &[u64],
    blank: u64,
    similarity: &dyn Fn(usize, usize) -> f64,
) -> Alignment {
    let mut a = Alignment::default();
    let mut unpaired_old = Vec::new();
    let mut unpaired_new = Vec::new();
    let pair = |a: &mut Alignment, o: usize, n: usize| {
        a.old_to_new.insert(o as u32 + 1, n as u32 + 1);
        a.new_to_old.insert(n as u32 + 1, o as u32 + 1);
    };
    for op in capture_diff_slices(Algorithm::Myers, old, new) {
        match op {
            DiffOp::Equal {
                old_index,
                new_index,
                len,
            } => {
                for k in 0..len {
                    pair(&mut a, old_index + k, new_index + k);
                }
            }
            DiffOp::Delete {
                old_index, old_len, ..
            } => unpaired_old.extend(old_index..old_index + old_len),
            DiffOp::Insert {
                new_index, new_len, ..
            } => unpaired_new.extend(new_index..new_index + new_len),
            DiffOp::Replace {
                old_index,
                old_len,
                new_index,
                new_len,
            } => {
                // As many rows replaced as replaced them: the same rows, edited in place. Only
                // when the counts differ is similarity needed to say which rows are which.
                let pairs = if old_len == new_len {
                    (0..old_len)
                        .map(|k| (old_index + k, new_index + k))
                        .collect()
                } else {
                    pair_block(
                        old_index..old_index + old_len,
                        new_index..new_index + new_len,
                        similarity,
                    )
                };
                let mut used_old = vec![false; old_len];
                let mut used_new = vec![false; new_len];
                for (o, n) in pairs {
                    pair(&mut a, o, n);
                    used_old[o - old_index] = true;
                    used_new[n - new_index] = true;
                }
                unpaired_old.extend(
                    (0..old_len)
                        .filter(|&k| !used_old[k])
                        .map(|k| old_index + k),
                );
                unpaired_new.extend(
                    (0..new_len)
                        .filter(|&k| !used_new[k])
                        .map(|k| new_index + k),
                );
            }
        }
    }
    // Moves: identical non-blank content on both sides.
    let mut by_sig: HashMap<u64, Vec<usize>> = HashMap::new();
    for &n in &unpaired_new {
        if new[n] != blank {
            by_sig.entry(new[n]).or_default().push(n);
        }
    }
    let mut moved_new = Vec::new();
    for &o in &unpaired_old {
        if old[o] == blank {
            continue;
        }
        if let Some(cands) = by_sig.get_mut(&old[o])
            && !cands.is_empty()
        {
            let n = cands.remove(0);
            pair(&mut a, o, n);
            a.moved.push((o as u32 + 1, n as u32 + 1));
            moved_new.push(n);
        }
    }
    let moved_old: Vec<usize> = a.moved.iter().map(|m| m.0 as usize - 1).collect();
    a.deleted = unpaired_old
        .into_iter()
        .filter(|o| !moved_old.contains(o))
        .map(|o| o as u32 + 1)
        .collect();
    a.inserted = unpaired_new
        .into_iter()
        .filter(|n| !moved_new.contains(n))
        .map(|n| n as u32 + 1)
        .collect();
    a.deleted.sort_unstable();
    a.inserted.sort_unstable();
    a
}

fn pair_block(
    old: std::ops::Range<usize>,
    new: std::ops::Range<usize>,
    similarity: &dyn Fn(usize, usize) -> f64,
) -> Vec<(usize, usize)> {
    let o: Vec<usize> = old.collect();
    let n: Vec<usize> = new.collect();
    pair_sorted(&o, &n, similarity)
}

/// Pairs entries of two ascending index lists by similarity without crossing: a greedy walk that
/// takes, for each old entry, the most similar new entry not before the last one taken.
fn pair_sorted(
    old: &[usize],
    new: &[usize],
    similarity: &dyn Fn(usize, usize) -> f64,
) -> Vec<(usize, usize)> {
    let mut pairs = Vec::new();
    let mut next_new = 0;
    for &o in old {
        let mut best: Option<(usize, f64)> = None;
        for (k, &n) in new.iter().enumerate().skip(next_new) {
            let s = similarity(o, n);
            if s >= PAIR_THRESHOLD && best.is_none_or(|(_, b)| s > b) {
                best = Some((k, s));
            }
        }
        if let Some((k, _)) = best {
            pairs.push((o, new[k]));
            next_new = k + 1;
        }
    }
    pairs
}

#[cfg(test)]
mod tests {
    use super::*;

    fn run(old: &[u64], new: &[u64]) -> Alignment {
        align(old, new, 0, &|o, n| {
            if old[o] / 10 == new[n] / 10 { 0.9 } else { 0.0 }
        })
    }

    #[test]
    fn insert_shifts_and_edit_pairs() {
        // 20 is edited to 21; 99 is inserted.
        let a = run(&[10, 20, 30], &[10, 99, 21, 30]);
        assert_eq!(a.inserted, vec![2]);
        assert!(a.deleted.is_empty());
        assert_eq!(a.old_to_new.get(&2), Some(&3));
        assert_eq!(a.exact(3), Some(4));
        assert_eq!(a.exact(7), Some(8));
    }

    #[test]
    fn move_is_detected() {
        let a = run(&[10, 20, 30, 40], &[10, 30, 40, 20]);
        assert_eq!(a.moved.len(), 1);
        assert!(a.inserted.is_empty() && a.deleted.is_empty());
    }
}
