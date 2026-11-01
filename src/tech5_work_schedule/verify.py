"""Verify a generated workbook against the config that produced it.

This replaces the original ``inspect_totals.py``, which hardcoded the
rough draft's layout, found no Total column in it, and exited 0 anyway.
``verify_workbook`` returns every mismatch it finds; an empty list means
the workbook matches the config exactly.

Since v2.1.0 this checks content *and* key styling, both against the
same ``Theme`` the renderer uses: titles, ids, descriptions, markers,
Total texts, legend, and notes, plus header/phase/deliverable fills,
milestone marker colour and boldness, the font name, freeze panes, and
print setup (orientation, paper size, fit-to-width, print titles). It
does not check every cell's font or border — a workbook hand-edited in
a spreadsheet app will usually differ in some un-checked detail, and
that is fine. Pass a custom Theme when verifying a workbook rendered
with one; a workbook rendered with a custom theme fails against the
default theme, by design.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any
from zipfile import BadZipFile

import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

from .config import Config, load_config
from .errors import ConfigError
from .layout import (
    FIRST_DATA_ROW,
    HEADER_ROW_1,
    HEADER_ROW_2,
    SUBTITLE_ROW,
    TITLE_ROW,
    Layout,
)
from .markers import ACTIVITY_MARKER, MILESTONE_MARKER, PHASE_MARKER, marker_for
from .render import SHEET_TITLE
from .theme import DEFAULT_THEME, Theme, argb

_MARKERS = {PHASE_MARKER, ACTIVITY_MARKER, MILESTONE_MARKER}


def _fill_rgb(cell: Any) -> str | None:
    """The ARGB colour of a cell's fill, or None if it has no fill."""
    color = cell.fill.start_color if cell.fill is not None else None
    return color.rgb if color is not None else None


def _normalise_title_rows(value: str | None) -> str:
    """Normalise print titles for comparison ("$1:$4" == "1:4")."""
    return (value or "").replace("$", "")


def verify_workbook(
    workbook_path: str | Path, config: Config, theme: Theme = DEFAULT_THEME
) -> list[str]:
    """Check a workbook's content and key styling against a config.

    Returns [] on a match. ``theme`` must be the theme the workbook was
    rendered with (the default for workbooks from ``build_workbook``).
    """
    mismatches: list[str] = []
    path = Path(workbook_path)
    if not path.exists():
        return [f"workbook not found: {path}"]
    try:
        wb = openpyxl.load_workbook(path)
    except (OSError, InvalidFileException, BadZipFile) as exc:
        return [f"cannot read workbook {path}: {exc}"]
    if SHEET_TITLE not in wb.sheetnames:
        return [f"missing sheet {SHEET_TITLE!r}; sheets are {wb.sheetnames}"]
    ws = wb[SHEET_TITLE]

    layout = Layout(config.meta.total_duration)
    total_col = layout.total_col

    title = ws.cell(row=TITLE_ROW, column=1).value
    if title != config.meta.project_name:
        mismatches.append(
            f"title row: expected {config.meta.project_name!r}, got {title!r}"
        )
    subtitle = ws.cell(row=SUBTITLE_ROW, column=1).value or ""
    expected_duration = (
        f"Duration: {config.meta.total_duration} {config.meta.time_unit_label}"
    )
    if expected_duration not in str(subtitle):
        mismatches.append(
            f"subtitle row: expected {expected_duration!r} in {subtitle!r}"
        )
    if config.meta.rfp_reference and config.meta.rfp_reference not in str(subtitle):
        mismatches.append(
            f"subtitle row: expected RFP reference {config.meta.rfp_reference!r}"
        )

    header = ws.cell(row=HEADER_ROW_2, column=total_col).value
    if header != "Total":
        msg = f"expected 'Total' in column {total_col}, got {header!r}"
        mismatches.append(f"header row {HEADER_ROW_2}: {msg}")
    for i in range(1, config.meta.total_duration + 1):
        value = ws.cell(row=HEADER_ROW_2, column=2 + i).value
        if value != i:
            msg = f"expected period {i} in column {2 + i}, got {value!r}"
            mismatches.append(f"header row {HEADER_ROW_2}: {msg}")

    for offset, item in enumerate(config.items):
        row = layout.data_row(offset)
        item_id = ws.cell(row=row, column=1).value
        if item_id != item.id:
            mismatches.append(f"row {row}: expected id {item.id!r}, got {item_id!r}")
        description = ws.cell(row=row, column=2).value
        if description != item.description:
            msg = f"expected description {item.description!r}, got {description!r}"
            mismatches.append(f"row {row}: {msg}")
        total_value = ws.cell(row=row, column=total_col).value
        if total_value != item.total_text:
            mismatches.append(
                f"row {row}: expected Total {item.total_text!r}, got {total_value!r}"
            )
        marker = marker_for(item.type, item.is_milestone)
        for period in range(1, config.meta.total_duration + 1):
            cell = ws.cell(row=row, column=2 + period)
            expected = marker.char if period in item.active_periods else None
            if cell.value != expected:
                mismatches.append(
                    f"row {row} period {period}: expected {expected!r}, "
                    f"got {cell.value!r}"
                )
            if expected is not None:
                if cell.font.bold != marker.bold:
                    mismatches.append(
                        f"row {row} period {period}: expected bold "
                        f"{marker.bold}, got {cell.font.bold}"
                    )
                if marker.color is not None and (
                    cell.font.color is None or cell.font.color.rgb != argb(marker.color)
                ):
                    mismatches.append(
                        f"row {row} period {period}: expected marker colour "
                        f"{argb(marker.color)}, got "
                        f"{cell.font.color.rgb if cell.font.color else None}"
                    )
        for col in range(1, total_col + 1):
            value = ws.cell(row=row, column=col).value
            if value in _MARKERS and col in (1, 2, total_col):
                msg = f"marker {value!r} in column {col} (id/description/total only)"
                mismatches.append(f"row {row}: {msg}")

    # Key styling, checked against the theme the workbook was rendered
    # with. A style-stripped workbook fails here, which is the point:
    # v2.0.1's content-only check could not tell it apart from a real
    # generated workbook.
    header_fill = _fill_rgb(ws.cell(row=HEADER_ROW_1, column=1))
    if header_fill != argb(theme.header_fill):
        mismatches.append(
            f"header fill: expected {argb(theme.header_fill)}, got {header_fill}"
        )
    for offset, item in enumerate(config.items):
        row = layout.data_row(offset)
        expected_fill = None
        if item.type == "Phase":
            expected_fill = theme.phase_fill
        elif item.type == "Deliverable":
            expected_fill = theme.deliverable_fill
        if expected_fill is not None:
            actual = _fill_rgb(ws.cell(row=row, column=1))
            if actual != argb(expected_fill):
                mismatches.append(
                    f"row {row} fill: expected {argb(expected_fill)}, got {actual}"
                )
    title_font = ws.cell(row=TITLE_ROW, column=1).font
    if title_font.name != theme.font_name:
        mismatches.append(
            f"title font: expected {theme.font_name!r}, got {title_font.name!r}"
        )
    if ws.freeze_panes != layout.freeze_panes:
        mismatches.append(
            f"freeze panes: expected {layout.freeze_panes!r}, got {ws.freeze_panes!r}"
        )
    page_setup = ws.page_setup
    if page_setup.orientation != "landscape":
        mismatches.append(
            f"page orientation: expected 'landscape', got {page_setup.orientation!r}"
        )
    # openpyxl's PAPERSIZE_A4 is the string "9" while a loaded workbook
    # reports the integer 9; compare numerically to avoid a false alarm.
    if int(page_setup.paperSize or 0) != int(ws.PAPERSIZE_A4):
        mismatches.append(
            f"paper size: expected A4 ({ws.PAPERSIZE_A4}), got {page_setup.paperSize}"
        )
    if page_setup.fitToWidth != 1:
        mismatches.append(f"fit to width: expected 1, got {page_setup.fitToWidth}")
    if _normalise_title_rows(ws.print_title_rows) != _normalise_title_rows(
        layout.print_title_rows
    ):
        mismatches.append(
            f"print titles: expected {layout.print_title_rows!r}, "
            f"got {ws.print_title_rows!r}"
        )

    # Legend and notes must be present somewhere below the data.
    texts = [
        str(cell.value)
        for row in ws.iter_rows(min_row=FIRST_DATA_ROW + len(config.items))
        for cell in row
        if cell.value is not None
    ]
    joined = "\n".join(texts)
    if "Legend:" not in joined:
        mismatches.append("missing 'Legend:' below the schedule")
    if (
        MILESTONE_MARKER not in joined
        or "Deliverable / Benchmark / Approval" not in joined
    ):
        mismatches.append("missing milestone legend entry")
    if "Duration of activities shall be indicated" not in joined:
        mismatches.append("missing World Bank guidance notes")
    return mismatches


def verify_files(
    workbook_path: str | Path,
    config_path: str | Path,
    theme: Theme = DEFAULT_THEME,
) -> list[str]:
    """Load both files and verify the workbook against the config."""
    try:
        config = load_config(config_path)
    except ConfigError as exc:
        return [str(exc)]
    return verify_workbook(workbook_path, config, theme=theme)
