"""Marker rules for the TECH-5 bar chart.

This module is the single source of truth for which character a schedule
item renders as. The rule is the legend's contract, not a rendering
detail: phases are █, activities are X, and deliverables — plus any
item flagged as a milestone — are a red ♦. Previously the rule lived in
a private generator helper and was duplicated (with its own copy of the
logic) in the verifier and the tests, so the three could drift apart.
Now the renderer, the verifier, and the tests all import this module,
and any future item type is added here once, with its legend entry.
"""

from __future__ import annotations

from dataclasses import dataclass

PHASE_MARKER = "█"
ACTIVITY_MARKER = "X"
MILESTONE_MARKER = "♦"


@dataclass(frozen=True)
class Marker:
    """A rendered marker: the character plus its font styling."""

    char: str
    bold: bool = False
    color: str | None = None
    """Hex colour (e.g. "FF0000"), or None for the default font colour."""


def marker_for(item_type: str, is_milestone: bool) -> Marker:
    """Return the marker for an item type and milestone flag.

    Deliverables always render as ♦, milestone or not, because the
    legend defines ♦ as "Deliverable / Benchmark / Approval". A
    milestone flag upgrades an Activity or Sub-Activity to ♦; it does
    not change a Phase, which is already a distinct block. Unknown types
    render as X, matching the old fallback — validation rejects them
    before rendering, so this branch is unreachable for valid configs.
    """
    if is_milestone or item_type == "Deliverable":
        return Marker(MILESTONE_MARKER, bold=True, color="FF0000")
    if item_type == "Phase":
        return Marker(PHASE_MARKER)
    return Marker(ACTIVITY_MARKER)
