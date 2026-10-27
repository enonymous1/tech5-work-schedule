import copy
import json
from pathlib import Path

import pytest

from tech5_work_schedule.config import ConfigError, load_config, parse_config

EXAMPLES = Path(__file__).resolve().parents[1] / "examples"


def _base() -> dict:
    return {
        "project_meta": {
            "project_name": "Example",
            "total_duration": 4,
            "time_unit_label": "Months",
        },
        "schedule_items": [
            {
                "id": "A-1",
                "description": "Phase one",
                "type": "Phase",
                "active_periods": [1],
            },
            {
                "id": "A-1.1",
                "description": "Activity",
                "type": "Activity",
                "active_periods": ["2-3"],
            },
        ],
    }


def test_sample_and_minimal_examples_load() -> None:
    for name in ("sample_tech5.json", "minimal_tech5.json"):
        config = load_config(EXAMPLES / name)
        assert config.items
        assert config.meta.total_duration >= 1


def test_sample_has_23_items_matching_original_structure() -> None:
    config = load_config(EXAMPLES / "sample_tech5.json")
    assert len(config.items) == 23
    assert [item.id for item in config.items][:5] == [
        "A-1",
        "A-1.1",
        "A-1.2",
        "D-1",
        "A-2",
    ]
    assert config.meta.rfp_reference == "{Insert RFP reference}"


def test_range_string_expands_and_dedupes() -> None:
    data = _base()
    data["schedule_items"][1]["active_periods"] = ["2-3", 3, "2"]
    config = parse_config(data)
    assert config.items[1].active_periods == (2, 3)


def test_old_sample_schema_is_rejected_with_guidance() -> None:
    data = {
        "project_metadata": {"name": "X"},
        "phases": [{"id": "A-1", "months": [1], "activities": []}],
    }
    with pytest.raises(ConfigError, match="project_meta.*schedule_items"):
        parse_config(data)


@pytest.mark.parametrize(
    "periods",
    [[0], [-1], [5], ["3-1"], ["bogus"], ["1-2-3"], [1.5], [True], [""]],
)
def test_invalid_periods_are_rejected(periods: list) -> None:
    data = _base()
    data["schedule_items"][0]["active_periods"] = periods
    with pytest.raises(ConfigError, match="active_periods"):
        parse_config(data)


def test_duplicate_ids_unknown_type_and_unknown_keys_rejected() -> None:
    data = _base()
    data["schedule_items"][1]["id"] = "A-1"
    with pytest.raises(ConfigError, match="duplicate id"):
        parse_config(data)

    data = _base()
    data["schedule_items"][0]["type"] = "Milestone"
    with pytest.raises(ConfigError, match="must be one of"):
        parse_config(data)

    data = _base()
    data["schedule_items"][0]["surprise"] = 1
    with pytest.raises(ConfigError, match="unknown key"):
        parse_config(data)

    data = _base()
    data["schedule_items"] = []
    with pytest.raises(ConfigError, match="non-empty"):
        parse_config(data)


def test_missing_file_and_bad_json_are_config_errors(tmp_path: Path) -> None:
    with pytest.raises(ConfigError, match="not found"):
        load_config(tmp_path / "missing.json")
    bad = tmp_path / "bad.json"
    bad.write_text("{not json", encoding="utf-8")
    with pytest.raises(ConfigError, match="invalid JSON"):
        load_config(bad)


def test_total_duration_must_be_positive_int() -> None:
    data = _base()
    data["project_meta"]["total_duration"] = 0
    with pytest.raises(ConfigError):
        parse_config(data)
    data = _base()
    data["project_meta"]["total_duration"] = True
    with pytest.raises(ConfigError):
        parse_config(data)


def test_config_is_not_mutated_by_parsing() -> None:
    data = _base()
    snapshot = copy.deepcopy(data)
    parse_config(data)
    assert data == snapshot
    assert json.dumps(data)
