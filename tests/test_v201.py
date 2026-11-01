from pathlib import Path

import pytest

from tech5_work_schedule.cli import main
from tech5_work_schedule.config import (
    MAX_TOTAL_DURATION,
    ConfigError,
    load_config,
    parse_config,
)
from tech5_work_schedule.generator import build_workbook, generate

EXAMPLES = Path(__file__).resolve().parents[1] / "examples"


def _base() -> dict:
    return {
        "project_meta": {"project_name": "Example", "total_duration": 4},
        "schedule_items": [
            {
                "id": "A-1",
                "description": "Phase one",
                "type": "Phase",
                "active_periods": [1],
            }
        ],
    }


def test_total_duration_cap() -> None:
    data = _base()
    data["project_meta"]["total_duration"] = MAX_TOTAL_DURATION
    assert parse_config(data).meta.total_duration == MAX_TOTAL_DURATION
    data["project_meta"]["total_duration"] = MAX_TOTAL_DURATION + 1
    with pytest.raises(ConfigError, match="must not exceed"):
        parse_config(data)


def test_duplicate_json_keys_rejected(tmp_path: Path) -> None:
    dup = tmp_path / "dup.json"
    dup.write_text(
        '{"project_meta": {"project_name": "First", "total_duration": 4}, '
        '"project_meta": {"project_name": "Second", "total_duration": 4}, '
        '"schedule_items": [{"id": "A-1", "description": "P", '
        '"type": "Phase", "active_periods": [1]}]}',
        encoding="utf-8",
    )
    with pytest.raises(ConfigError, match="duplicate key"):
        load_config(dup)


def test_guidance_note_2_has_single_period() -> None:
    config = load_config(EXAMPLES / "sample_tech5.json")
    ws = build_workbook(config)["TECH-5"]
    note2 = str(ws.cell(row=35, column=2).value)
    assert note2 == (
        "2 Duration of activities shall be indicated in a form of a bar chart."
    )
    assert ".." not in note2


def test_default_output_never_overwrites(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.chdir(tmp_path)
    first = generate(EXAMPLES / "minimal_tech5.json")
    second = generate(EXAMPLES / "minimal_tech5.json")
    assert first != second
    assert first.exists() and second.exists()
    assert second.name.endswith("-2.xlsx")


def test_cli_generate_output_is_directory_returns_2(tmp_path: Path, capsys) -> None:
    assert (
        main(
            [
                "generate",
                "--config",
                str(EXAMPLES / "minimal_tech5.json"),
                "--output",
                str(tmp_path),
            ]
        )
        == 2
    )
    err = capsys.readouterr().err
    assert "error: cannot write workbook:" in err
    assert "Traceback" not in err
