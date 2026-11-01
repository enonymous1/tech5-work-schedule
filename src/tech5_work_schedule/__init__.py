"""tech5-work-schedule: generate TECH-5 work-schedule workbooks from JSON."""

from __future__ import annotations

from .config import (
    Config,
    ItemType,
    ProjectMeta,
    ScheduleItem,
    load_config,
    parse_config,
)
from .errors import ConfigError
from .io import generate
from .layout import Layout
from .markers import Marker, marker_for
from .render import build_workbook
from .theme import DEFAULT_THEME, Theme
from .verify import verify_files, verify_workbook

__version__ = "2.1.0"

__all__ = [
    "DEFAULT_THEME",
    "Config",
    "ConfigError",
    "ItemType",
    "Layout",
    "Marker",
    "ProjectMeta",
    "ScheduleItem",
    "Theme",
    "__version__",
    "build_workbook",
    "generate",
    "load_config",
    "marker_for",
    "parse_config",
    "verify_files",
    "verify_workbook",
]
