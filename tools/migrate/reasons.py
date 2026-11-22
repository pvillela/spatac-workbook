"""Reasons for change, inferred from the old Working notes and the amounts (written into
Expenditures column U by migrate.py).

Keyed by item number, with a prefix of the building element (or GL account name, for Spatac
rows) where the item number alone is ambiguous. Each entry is (reason, doubt). A doubt becomes a
comment on the reason cell, for SPATAC to resolve before the plan goes to the Board.
"""

NO_NOTE = "No working note: reason inferred from the amounts (Actual to date). Please confirm."
DROPPED = (
    "No working note explains why the RFS amount was removed for 2026. Please confirm "
    "(deferred, not needed, or covered elsewhere?)."
)

REASONS = {
    ("2.2", ""): ("North patio project carried out in 2026; year-end projection.", None),
    ("3.4a", ""): ("Slab repairs in 2026 (year-end projection); not in the RFS this year.", None),
    ("3.9", ""): (
        "Moved from the 2025 RFS to 2028, inflation-adjusted.",
        "The note says 2027, but the November estimate puts the amount in 2028 (2027 is now "
        "zero). Please confirm the year.",
    ),
    ("3.10a", ""): ("Garage repairs in 2026 (year-end projection); not in the RFS this year.", None),
    ("4.12", ""): ("New gate arm per quote.", None),
    ("4.13", ""): ("Lower than the RFS: year-end projection.", None),
    ("5.1", ""): (
        "Not needed in 2026; window repairs tracked under 5.1a and 5.3.",
        DROPPED,
    ),
    ("5.1a", ""): (
        "Sealed-unit repairs (Hi-Lon) in 2026, year-end projection; not in the RFS this year.",
        None,
    ),
    ("5.3", ""): (
        "Glass repair booked here per the year-end projection; not in the RFS this year.",
        None,
    ),
    ("5.3.1b", ""): ("Door repairs in 2026 (actual); not in the RFS this year.", NO_NOTE),
    ("5.5", ""): ("Not expected in 2026.", DROPPED),
    ("5.6", ""): ("Not expected in 2026.", DROPPED),
    ("5.6a", ""): (
        "Balcony repairs as needed (T&T handle replacements); year-end projection; not in the "
        "RFS this year.",
        None,
    ),
    ("6.1.1a", ""): ("Consulting in 2026 (year-end projection); not in the RFS this year.", None),
    ("6.2", ""): ("Leak repairs in 2026 (year-end projection); not in the RFS this year.", None),
    ("6.3", ""): (
        "Moved from 2026 to 2028, inflation-adjusted.",
        "The $50,000 in 2027 is not explained in the notes. Please confirm.",
    ),
    ("7.1.1", ""): (
        "Garage door repair (Atlas sensors), 2026 actual; not in the RFS this year.",
        None,
    ),
    ("7.2a", ""): (
        "Garage door circuit board (Atlas), 2026 actual; not in the RFS this year.",
        None,
    ),
    ("10.10", ""): ("Lower than the RFS: year-end projection.", None),
    ("10.11", ""): (
        "Dry sprinkler work in 2026 (year-end projection); not in the RFS this year.",
        None,
    ),
    ("11.2", ""): (
        "Net credit in 2026: apparent insurance recovery.",
        "An earlier note assumed insurance would not pay. Please confirm the recovery.",
    ),
    ("11.5", ""): (
        "Not in the year-end projection; leak-investigation costs moved to 11.14a.",
        None,
    ),
    ("11.6", ""): (
        "Heat-pump repairs (motor, cover refurbishment), 2026 actual; not in the RFS this year.",
        None,
    ),
    ("11.9", ""): (
        "Glycol top-up and filter housing, 2026 actual; not in the RFS this year.",
        None,
    ),
    ("11.14a", ""): ("Heat-pump loop leak investigation (moved from 11.5), 2026 actual.", None),
    ("11.15", ""): ("BAS work in 2026 (actual); not in the RFS this year.", NO_NOTE),
    ("12.2a", ""): (
        "Make-up air unit condensate switch, 2026 actual; not in the RFS this year.",
        None,
    ),
    ("12.4", ""): (
        "Garage exhaust fan repairs after a power surge; year-end projection.",
        None,
    ),
    ("13.1", ""): ("Pilot replacement after the storm flood (moved from 11.2).", None),
    ("13.5", ""): (
        "Recirculation valve and piping work; year-end projection; not in the RFS this year.",
        None,
    ),
    ("13.8", ""): ("Sump and sanitary pump replacements, 2026 actual.", None),
    ("13.9", ""): (
        "Drain repairs and P1 preventive drain replacements, 2026 actual.",
        None,
    ),
    ("13.13", ""): ("Valve work in 2026 (actual); not in the RFS this year.", NO_NOTE),
    ("13.14", ""): ("Irrigation repairs, 2026 actual; not in the RFS this year.", None),
    ("14.11", ""): ("Exterior lighting repairs, 2026 actual.", NO_NOTE),
    ("15.5a", ""): (
        "Elevator modernization (with safeties) completed in 2026; the remaining placeholder, "
        "including the TK buy-out, dropped.",
        "The notes give two figures: 220,326 (actual plus pending Solucore, revised by Bruce, "
        "now the 2026 Late-year estimate) and the year-end projection 379,893 (including a December "
        "forecast of 154,085). Please confirm which applies.",
    ),
    ("16.1.1a", ""): ("Fire damper in the garbage chute, 2026 actual.", None),
    ("16.2a", ""): ("Bank charges booked to the reserve fund.", None),
    ("16.3", ""): ("Garbage bin replacement, 2026 actual.", None),
    ("16.4", ""): ("Garbage chute repairs, 2026 actual.", NO_NOTE),
    ("17.3", ""): ("Access control repairs, 2026 actual.", NO_NOTE),
    ("17.5", ""): (
        "Cell-phone booster replacement booked here (no dedicated line), 2026 actual.",
        None,
    ),
    ("18.1", ""): ("Carried over from the 2025 RFS; year-end projection.", None),
    ("18.3.1", "Refurbish Corridors"): (
        "Corridor refurbishment brought forward from 2028; P1 refurbishment (Tri-Can) complete "
        "in 2026.",
        "The notes mention bringing in only 1/7 of the 2028 amount. The $600,000 in 2027, and "
        "the removal of the whole 2028 amount, are not explained. Please confirm.",
    ),
    ("18.4", ""): ("Not expected in 2026.", DROPPED),
    ("18.8", ""): ("Lower than the RFS: year-end projection.", None),
    ("18.9", ""): (
        "Carried over from the 2025 RFS; the main cost is within the P1 contract (18.3.1).",
        None,
    ),
    ("18.10", ""): (
        "Carried over from the 2025 RFS; the main cost is within the P1 contract (18.3.1).",
        None,
    ),
    ("18.11", ""): (
        "Removed: not in the year-end projection.",
        "The note says 'Bring $42,448 from 2027', yet no amount remains in any year. Please "
        "confirm whether it was dropped or should appear in 2026.",
    ),
    ("18.14", ""): (
        "Guest suite work in 2026 (year-end projection); not in the RFS this year.",
        None,
    ),
    ("18.21a", ""): (
        "Pool and whirlpool repairs; year-end projection; not in the RFS this year.",
        None,
    ),
    ("18.22", ""): ("Power door opener repairs, 2026 actual.", NO_NOTE),
    ("18.23a", ""): ("Stairwell door closer, 2026 actual.", None),
    ("20.1", ""): (
        "Lower than the RFS: year-end projection 9,119, which excludes the storm-flood repair "
        "(possible insurance claim).",
        None,
    ),
    ("12.6a", ""): (
        "In-suite heat-pump repairs (condensate clog, filter), 2026 actual.",
        None,
    ),
}


def reason_for(item, label):
    """The (reason, doubt) for a row, or None."""
    exact = REASONS.get((item, label))
    if exact:
        return exact
    for (it, prefix), value in REASONS.items():
        if it == item and prefix and label.startswith(prefix):
            return value
    return REASONS.get((item, ""))
