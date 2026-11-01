"""Workbook styling, as data instead of scattered constants.

Previously fills, fonts, and column widths were module constants in the
generator, so a different look meant editing rendering code, and the
verifier could not check styling at all (it had nothing to compare
against). A ``Theme`` keeps every style decision in one frozen
dataclass: ``build_workbook(config, theme=...)`` renders with it, and
``verify_workbook(path, config, theme=...)`` checks the workbook
against it. The default theme reproduces the v2.0.0 look exactly, so
existing workbooks verify unchanged.
"""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl.styles import PatternFill


def argb(hex_color: str) -> str:
    """Normalise a 6-digit hex colour to openpyxl's 8-digit ARGB form."""
    return f"00{hex_color.upper()}"


@dataclass(frozen=True)
class Theme:
    """Colours, fonts, and column widths for a TECH-5 workbook."""

    font_name: str = "Gill Sans"
    font_size: int = 10
    title_size: int = 12
    header_fill: str = "BDD7EE"
    phase_fill: str = "DDEBF7"
    deliverable_fill: str = "E2EFDA"
    milestone_color: str = "FF0000"
    id_col_width: float = 8
    month_col_width: float = 3.5
    desc_col_min_width: float = 40
    desc_col_max_width: float = 120
    total_col_min_width: float = 10

    def fill(self, hex_color: str) -> PatternFill:
        return PatternFill(
            start_color=hex_color, end_color=hex_color, fill_type="solid"
        )


DEFAULT_THEME = Theme()
