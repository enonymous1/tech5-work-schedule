from pathlib import Path

import openpyxl
import pytest

from tech5_work_schedule import (
    ItemType,
    Layout,
    Theme,
    build_workbook,
    marker_for,
)
from tech5_work_schedule.config import ConfigError, load_config, parse_config
from tech5_work_schedule.generator import FIRST_DATA_ROW, generate, slugify
from tech5_work_schedule.io import generate as io_generate
from tech5_work_schedule.markers import ACTIVITY_MARKER, MILESTONE_MARKER, PHASE_MARKER
from tech5_work_schedule.verify import verify_files, verify_workbook

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


def test_marker_for_is_single_source_of_truth() -> None:
    assert marker_for("Phase", False).char == PHASE_MARKER
    assert marker_for("Activity", False).char == ACTIVITY_MARKER
    assert marker_for("Activity", True).char == MILESTONE_MARKER
    assert marker_for("Deliverable", False).char == MILESTONE_MARKER
    assert marker_for("Deliverable", False).bold is True
    assert marker_for("Deliverable", False).color == "FF0000"
    assert marker_for("Phase", False).bold is False


def test_layout_maths() -> None:
    layout = Layout(9)
    assert layout.total_col == 12
    assert layout.month_col == 3
    assert layout.data_row(0) == FIRST_DATA_ROW
    assert layout.last_data_row(23) == FIRST_DATA_ROW + 22
    assert layout.freeze_panes == "C5"
    assert layout.print_title_rows == "1:4"


def test_item_type_is_str_enum() -> None:
    config = load_config(EXAMPLES / "minimal_tech5.json")
    assert config.items[0].type is ItemType.PHASE
    # StrEnum compares equal to its plain string, so old code keeps working.
    assert config.items[0].type == "Phase"
    data = _base()
    data["schedule_items"][0]["type"] = "Bogus"
    with pytest.raises(ConfigError, match="must be one of"):
        parse_config(data)


def test_schema_version() -> None:
    assert load_config(EXAMPLES / "sample_tech5.json").schema_version == 1
    data = _base()
    data["schema_version"] = 1
    assert parse_config(data).schema_version == 1
    data["schema_version"] = 2
    with pytest.raises(ConfigError, match="schema_version"):
        parse_config(data)
    data["schema_version"] = True
    with pytest.raises(ConfigError, match="schema_version"):
        parse_config(data)


def test_custom_theme_changes_render_and_verify(tmp_path: Path) -> None:
    theme = Theme(header_fill="FF0000", font_name="Calibri")
    config = load_config(EXAMPLES / "minimal_tech5.json")
    wb = build_workbook(config, theme=theme)
    ws = wb["TECH-5"]
    assert ws.cell(row=3, column=1).fill.start_color.rgb == "00FF0000"
    assert ws.cell(row=1, column=1).font.name == "Calibri"
    path = tmp_path / "theme.xlsx"
    wb.save(path)
    assert verify_workbook(path, config, theme=theme) == []
    # The same workbook fails against the default theme, by design.
    assert verify_workbook(path, config) != []


def test_style_stripped_workbook_fails_verify(tmp_path: Path) -> None:
    workbook = generate(EXAMPLES / "minimal_tech5.json", tmp_path / "out.xlsx")
    wb = openpyxl.load_workbook(workbook)
    ws = wb["TECH-5"]
    for row in ws.iter_rows():
        for cell in row:
            cell.fill = openpyxl.styles.PatternFill(fill_type=None)
    wb.save(workbook)
    mismatches = verify_files(workbook, EXAMPLES / "minimal_tech5.json")
    assert any("fill" in m for m in mismatches)


def test_long_description_wraps_and_grows_row(tmp_path: Path) -> None:
    data = _base()
    data["schedule_items"][0]["description"] = "A very long description " * 6
    config = parse_config(data)
    ws = build_workbook(config)["TECH-5"]
    desc = ws.cell(row=FIRST_DATA_ROW, column=2)
    assert desc.alignment.wrap_text is True
    assert ws.row_dimensions[FIRST_DATA_ROW].height is not None
    assert ws.row_dimensions[FIRST_DATA_ROW].height > 15


def test_note_heights_grow_on_narrow_schedule() -> None:
    config = load_config(EXAMPLES / "minimal_tech5.json")
    ws = build_workbook(config)["TECH-5"]
    # Minimal config: last data row 9, legend label 11, legend items
    # 12-14, notes from 16. Note 1 is long and the 4-month table is
    # narrow, so it needs more than the old fixed 45 points.
    note1_height = ws.row_dimensions[16].height
    assert note1_height is not None and note1_height > 45


def test_slugify_keeps_unicode_and_removes_separators() -> None:
    assert slugify("Café — Phase 2") == "Café_Phase_2"
    assert slugify("A/B: Test") == "AB_Test"
    assert slugify("My Project: 2026!") == "My_Project_2026"
    assert slugify("   ") == "Output"
    assert len(slugify("x" * 100)) == 60


def test_io_generate_is_same_function(tmp_path: Path) -> None:
    assert io_generate(EXAMPLES / "minimal_tech5.json", tmp_path / "out.xlsx").exists()


def test_public_api_exports() -> None:
    import tech5_work_schedule as pkg

    for name in (
        "Config",
        "ItemType",
        "Layout",
        "Theme",
        "DEFAULT_THEME",
        "build_workbook",
        "generate",
        "load_config",
        "marker_for",
        "parse_config",
        "verify_files",
        "verify_workbook",
        "__version__",
    ):
        assert hasattr(pkg, name), name
    assert pkg.__version__ == "2.1.0"
