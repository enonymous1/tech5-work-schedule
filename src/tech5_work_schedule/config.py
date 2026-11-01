"""Load and validate TECH-5 JSON configs.

The single supported schema is::

    {
      "project_meta": {
        "project_name": "Example Schedule",
        "rfp_reference": "{Insert RFP reference}",
        "time_unit_label": "Months",
        "total_duration": 9
      },
      "schedule_items": [
        {
          "id": "A-1",
          "description": "PHASE 1: INCEPTION",
          "type": "Phase",
          "active_periods": [1],
          "is_milestone": false,
          "total_text": "2 Weeks"
        }
      ]
    }

Validation is strict and fail-loud: unknown keys, missing keys, duplicate
ids, and out-of-range periods are errors, never silently skipped. The
original project's sample config used a different schema
(``project_metadata`` / ``phases`` / ``months``) that the generator never
read, so it silently produced an empty schedule; that schema is rejected
here with a message that names the expected keys.
"""

from __future__ import annotations

import json
from collections.abc import Collection
from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path
from typing import Any

from .errors import ConfigError


class ItemType(StrEnum):
    """The four schedule-item kinds, as written in configs.

    A StrEnum (Python 3.11+) keeps the config-facing strings — "Phase",
    "Activity", "Deliverable", "Sub-Activity" — while giving the rest of
    the code a closed type instead of a bare ``str`` that accepts
    anything. Comparisons with plain strings still work, so existing
    code comparing ``item.type == "Phase"`` is unaffected.
    """

    PHASE = "Phase"
    ACTIVITY = "Activity"
    DELIVERABLE = "Deliverable"
    SUB_ACTIVITY = "Sub-Activity"


ITEM_TYPES = frozenset(str(t) for t in ItemType)

# The only schema version this release accepts. Versioning is cheap
# insurance: the next schema change can keep accepting 1 and add 2,
# instead of breaking every existing config at once.
SCHEMA_VERSION = 1

# A TECH-5 bar chart is unreadable under fit-to-width printing long
# before Excel's 16,384-column limit, so the schedule length is capped.
# Split longer programmes into phases or use a coarser time unit.
MAX_TOTAL_DURATION = 240

_META_KEYS = frozenset(
    {"project_name", "rfp_reference", "time_unit_label", "total_duration"}
)
_ITEM_KEYS = frozenset(
    {"id", "description", "type", "active_periods", "is_milestone", "total_text"}
)
_TOP_KEYS = frozenset({"project_meta", "schedule_items", "schema_version"})


@dataclass(frozen=True)
class ProjectMeta:
    """Project-level metadata: name, length, time-unit label, RFP reference.

    ``time_unit_label`` is a display label only — periods are ordinals
    in that unit, and no date arithmetic is performed anywhere.
    """

    project_name: str
    total_duration: int
    time_unit_label: str = "Months"
    rfp_reference: str | None = None


@dataclass(frozen=True)
class ScheduleItem:
    """One schedule row.

    ``total_text`` is author-supplied display text for the Total column
    (for example "2 Weeks" or "Week 10"), not a value computed from
    ``active_periods``. The renderer prints it verbatim, so a wrong
    total is rendered faithfully — check totals before generating.
    """

    id: str
    description: str
    type: ItemType
    active_periods: tuple[int, ...]
    is_milestone: bool = False
    total_text: str | None = None


@dataclass(frozen=True)
class Config:
    """A fully validated config: metadata plus its schedule items."""

    meta: ProjectMeta
    items: tuple[ScheduleItem, ...]
    schema_version: int = SCHEMA_VERSION


def _fail(path: str, message: str) -> ConfigError:
    return ConfigError(f"{path}: {message}")


def _check_keys(
    data: dict[str, Any],
    allowed: frozenset[str],
    required: Collection[str],
    path: str,
) -> None:
    unknown = sorted(set(data) - allowed)
    if unknown:
        raise _fail(path, f"unknown key(s): {', '.join(unknown)}")
    missing = sorted(set(required) - set(data))
    if missing:
        raise _fail(path, f"missing key(s): {', '.join(missing)}")


def _as_non_empty_str(value: Any, path: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise _fail(path, "must be a non-empty string")
    return value


def _as_int(value: Any, path: str) -> int:
    # bool is a subclass of int; a period of True/False is never valid.
    if isinstance(value, bool) or not isinstance(value, int):
        raise _fail(path, f"must be an integer, got {value!r}")
    return value


def _parse_period(value: Any, path: str, total_duration: int) -> list[int]:
    """Expand one active_period entry to concrete period numbers."""
    if isinstance(value, int) and not isinstance(value, bool):
        periods = [value]
    elif isinstance(value, str):
        text = value.strip()
        if "-" in text:
            parts = text.split("-")
            if len(parts) != 2:
                raise _fail(path, f"invalid range {value!r}; expected 'start-end'")
            try:
                start = int(parts[0].strip())
                end = int(parts[1].strip())
            except ValueError:
                raise _fail(path, f"invalid range {value!r}; bounds must be integers")
            if start > end:
                raise _fail(path, f"invalid range {value!r}; start must not exceed end")
            periods = list(range(start, end + 1))
        else:
            try:
                periods = [int(text)]
            except ValueError:
                raise _fail(
                    path,
                    f"invalid period {value!r}; expected an integer or 'start-end'",
                )
    else:
        raise _fail(path, f"must be an integer or 'start-end' string, got {value!r}")
    for period in periods:
        if period < 1 or period > total_duration:
            raise _fail(
                path,
                f"period {period} is outside 1..{total_duration}",
            )
    return periods


def _parse_meta(data: Any) -> ProjectMeta:
    """Validate project metadata, including the duration cap.

    The upper bound exists because the renderer creates one column per
    period: an unbounded duration would build a workbook Excel cannot
    open (see MAX_TOTAL_DURATION).
    """
    if not isinstance(data, dict):
        raise _fail("project_meta", "must be an object")
    _check_keys(data, _META_KEYS, {"project_name", "total_duration"}, "project_meta")
    total_duration = _as_int(data["total_duration"], "project_meta.total_duration")
    if total_duration < 1:
        raise _fail("project_meta.total_duration", "must be at least 1")
    if total_duration > MAX_TOTAL_DURATION:
        raise _fail(
            "project_meta.total_duration",
            f"must not exceed {MAX_TOTAL_DURATION}; split the schedule "
            "into phases or use a coarser time unit",
        )
    rfp_reference = data.get("rfp_reference")
    if rfp_reference is not None:
        rfp_reference = _as_non_empty_str(rfp_reference, "project_meta.rfp_reference")
    time_unit_label = data.get("time_unit_label", "Months")
    time_unit_label = _as_non_empty_str(time_unit_label, "project_meta.time_unit_label")
    return ProjectMeta(
        project_name=_as_non_empty_str(
            data["project_name"], "project_meta.project_name"
        ),
        total_duration=total_duration,
        time_unit_label=time_unit_label,
        rfp_reference=rfp_reference,
    )


def _parse_item(data: Any, index: int, total_duration: int) -> ScheduleItem:
    """Validate one schedule item, including its type and periods.

    Periods are expanded (ranges become concrete numbers), checked
    against the duration, and de-duplicated, so the renderer never has
    to reason about ranges or overlaps.
    """
    path = f"schedule_items[{index}]"
    if not isinstance(data, dict):
        raise _fail(path, "must be an object")
    _check_keys(data, _ITEM_KEYS, {"id", "description", "type", "active_periods"}, path)
    raw_type = _as_non_empty_str(data["type"], f"{path}.type")
    try:
        item_type = ItemType(raw_type)
    except ValueError:
        raise _fail(
            f"{path}.type",
            f"must be one of {', '.join(sorted(ITEM_TYPES))}, got {raw_type!r}",
        ) from None
    raw_periods = data["active_periods"]
    if not isinstance(raw_periods, list) or not raw_periods:
        raise _fail(f"{path}.active_periods", "must be a non-empty list")
    expanded: list[int] = []
    for i, raw in enumerate(raw_periods):
        expanded.extend(
            _parse_period(raw, f"{path}.active_periods[{i}]", total_duration)
        )
    # De-duplicate, keeping period order, so overlapping ranges render once.
    periods = tuple(dict.fromkeys(expanded))
    is_milestone = data.get("is_milestone", False)
    if not isinstance(is_milestone, bool):
        raise _fail(f"{path}.is_milestone", "must be true or false")
    total_text = data.get("total_text")
    if total_text is not None:
        total_text = _as_non_empty_str(total_text, f"{path}.total_text")
    return ScheduleItem(
        id=_as_non_empty_str(data["id"], f"{path}.id"),
        description=_as_non_empty_str(data["description"], f"{path}.description"),
        type=item_type,
        active_periods=periods,
        is_milestone=is_milestone,
        total_text=total_text,
    )


def parse_config(data: Any) -> Config:
    """Validate a decoded JSON document and return a Config.

    ``schema_version`` is optional and defaults to 1; any other value
    is rejected so a config written for a future schema fails loudly
    instead of being misread under today's rules.
    """
    if not isinstance(data, dict):
        raise ConfigError("config root: must be a JSON object")
    if "project_metadata" in data or "phases" in data:
        raise ConfigError(
            "config root: unsupported schema using 'project_metadata'/'phases'; "
            "use 'project_meta' and 'schedule_items' (see examples/sample_tech5.json)"
        )
    _check_keys(data, _TOP_KEYS, {"project_meta", "schedule_items"}, "config root")
    schema_version = data.get("schema_version", SCHEMA_VERSION)
    if (
        isinstance(schema_version, bool)
        or not isinstance(schema_version, int)
        or schema_version != SCHEMA_VERSION
    ):
        raise ConfigError(
            f"config root.schema_version: must be {SCHEMA_VERSION}, "
            f"got {schema_version!r}"
        )
    meta = _parse_meta(data["project_meta"])
    raw_items = data["schedule_items"]
    if not isinstance(raw_items, list) or not raw_items:
        raise ConfigError("schedule_items: must be a non-empty list")
    items = tuple(
        _parse_item(raw, index, meta.total_duration)
        for index, raw in enumerate(raw_items)
    )
    seen: set[str] = set()
    for item in items:
        if item.id in seen:
            raise ConfigError(f"schedule_items: duplicate id {item.id!r}")
        seen.add(item.id)
    return Config(meta=meta, items=items, schema_version=schema_version)


def _no_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    """``object_pairs_hook`` that rejects duplicate JSON keys.

    ``json.loads`` silently keeps the last value for a repeated key, so
    a copy-paste error in a config would disappear without a trace.
    Raising here (a ConfigError, not a JSONDecodeError) makes the
    duplicate a named validation failure instead.
    """
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ConfigError(f"duplicate key {key!r} in config JSON")
        result[key] = value
    return result


def load_config(path: str | Path) -> Config:
    """Read and validate a config file.

    Duplicate JSON keys are rejected (see ``_no_duplicates``); all other
    failures are reported as ConfigError with the file path included.
    """
    config_path = Path(path)
    try:
        text = config_path.read_text(encoding="utf-8")
    except FileNotFoundError:
        raise ConfigError(f"config file not found: {config_path}") from None
    except OSError as exc:
        raise ConfigError(f"cannot read config file {config_path}: {exc}") from exc
    try:
        data = json.loads(text, object_pairs_hook=_no_duplicates)
    except json.JSONDecodeError as exc:
        raise ConfigError(f"{config_path}: invalid JSON: {exc}") from exc
    try:
        return parse_config(data)
    except ConfigError as exc:
        raise ConfigError(f"{config_path}: {exc}") from exc
