"""Backwards-compatible re-exports for the v2.1.0 render/io split.

Rendering moved to ``render`` and file output to ``io``; this module
keeps the old import paths working (``from .generator import
build_workbook`` and friends) so existing code and tests do not need to
change. New code should import from ``render`` or ``io`` directly.
"""

from __future__ import annotations

from .io import generate, slugify
from .layout import (
    FIRST_DATA_ROW,
    HEADER_ROW_1,
    HEADER_ROW_2,
    SUBTITLE_ROW,
    TITLE_ROW,
)
from .markers import ACTIVITY_MARKER, MILESTONE_MARKER, PHASE_MARKER
from .render import SHEET_TITLE, THIN_BORDER, build_workbook

__all__ = [
    "ACTIVITY_MARKER",
    "FIRST_DATA_ROW",
    "HEADER_ROW_1",
    "HEADER_ROW_2",
    "MILESTONE_MARKER",
    "PHASE_MARKER",
    "SHEET_TITLE",
    "SUBTITLE_ROW",
    "THIN_BORDER",
    "TITLE_ROW",
    "build_workbook",
    "generate",
    "slugify",
]
