"""Verify a generated workbook against the config that produced it.

This replaces the original ``inspect_totals.py``, which hardcoded the
rough draft's layout, found no Total column in it, and exited 0 anyway.
``verify_workbook`` returns every mismatch it finds; an empty list means
the workbook matches the config exactly.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl

from .config import Config, load_config
from .errors import ConfigError
from .generator import (
    FIRST_DATA_ROW,
    HEADER_ROW_2,
    MILESTONE_MARKER,
    SHEET_TITLE,
    SUBTITLE_ROW,
    TITLE_ROW,
    _marker_for,
)

_MARKERS = {"█", "X", "♦"}


def verify_workbook(workbook_path: str | Path, config: Config) -> list[str]:
    mismatches: list[str] = []
    path = Path(workbook_path)
    if not path.exists():
        return [f"workbook not found: {path}"]
    try:
        wb = openpyxl.load_workbook(path)
    except Exception as exc:  # openpyxl raises several types for bad files
        return [f"cannot read workbook {path}: {exc}"]
    if SHEET_TITLE not in wb.sheetnames:
        return [f"missing sheet {SHEET_TITLE!r}; sheets are {wb.sheetnames}"]
    ws = wb[SHEET_TITLE]

    total_col = 3 + config.meta.total_duration

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
        row = FIRST_DATA_ROW + offset
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
        marker, _ = _marker_for(item.type, item.is_milestone)
        for period in range(1, config.meta.total_duration + 1):
            cell = ws.cell(row=row, column=2 + period).value
            expected = marker if period in item.active_periods else None
            if cell != expected:
                mismatches.append(
                    f"row {row} period {period}: expected {expected!r}, got {cell!r}"
                )
        for col in range(1, total_col + 1):
            value = ws.cell(row=row, column=col).value
            if value in _MARKERS and col in (1, 2, total_col):
                msg = f"marker {value!r} in column {col} (id/description/total only)"
                mismatches.append(f"row {row}: {msg}")

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


def verify_files(workbook_path: str | Path, config_path: str | Path) -> list[str]:
    """Load both files and verify the workbook against the config."""
    try:
        config = load_config(config_path)
    except ConfigError as exc:
        return [str(exc)]
    return verify_workbook(workbook_path, config)
