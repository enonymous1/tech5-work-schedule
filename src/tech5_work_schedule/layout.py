"""The TECH-5 sheet layout, computed once per workbook.

Previously every consumer recomputed row and column numbers from the
duration by hand — the renderer, the verifier, and the tests each
derived the Total column and the first data row independently. A layout
change then had to land in three places to stay consistent. Now a
``Layout`` is built from the duration alone and passed around; the
renderer and the verifier read the same numbers.
"""

from __future__ import annotations

from dataclasses import dataclass

TITLE_ROW = 1
SUBTITLE_ROW = 2
HEADER_ROW_1 = 3
HEADER_ROW_2 = 4
FIRST_DATA_ROW = 5


@dataclass(frozen=True)
class Layout:
    """Row and column numbers for a TECH-5 sheet of a given duration."""

    total_duration: int

    @property
    def total_col(self) -> int:
        """The Total column: id (1) + description (2) + one per period."""
        return 3 + self.total_duration

    @property
    def month_col(self) -> int:
        """The first period column."""
        return 3

    def data_row(self, index: int) -> int:
        """The sheet row for the ``index``-th item (0-based)."""
        return FIRST_DATA_ROW + index

    def last_data_row(self, item_count: int) -> int:
        return FIRST_DATA_ROW + item_count - 1

    @property
    def freeze_panes(self) -> str:
        """Freeze below the header and left of the period columns."""
        return f"C{FIRST_DATA_ROW}"

    @property
    def print_title_rows(self) -> str:
        return f"{TITLE_ROW}:{HEADER_ROW_2}"
