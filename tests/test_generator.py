from pathlib import Path

import openpyxl

from tech5_work_schedule.config import load_config
from tech5_work_schedule.generator import (
    FIRST_DATA_ROW,
    build_workbook,
    generate,
    slugify,
)

EXAMPLES = Path(__file__).resolve().parents[1] / "examples"


def _sample_sheet():
    config = load_config(EXAMPLES / "sample_tech5.json")
    wb = build_workbook(config)
    return config, wb["TECH-5"]


def test_title_subtitle_and_header_rows() -> None:
    config, ws = _sample_sheet()
    assert ws.cell(row=1, column=1).value == config.meta.project_name
    subtitle = ws.cell(row=2, column=1).value
    assert "{Insert RFP reference}" in subtitle
    assert "Duration: 9 Months" in subtitle
    assert ws.cell(row=4, column=12).value == "Total"
    assert [ws.cell(row=4, column=c).value for c in range(3, 12)] == list(range(1, 10))
    assert ws.freeze_panes == "C5"
    assert ws.print_title_rows == "$1:$4"


def test_marker_grid_matches_config_exactly() -> None:
    config, ws = _sample_sheet()
    for offset, item in enumerate(config.items):
        row = FIRST_DATA_ROW + offset
        assert ws.cell(row=row, column=1).value == item.id
        assert ws.cell(row=row, column=2).value == item.description
        expected_marker = (
            "♦"
            if item.type == "Deliverable" or item.is_milestone
            else ("█" if item.type == "Phase" else "X")
        )
        for period in range(1, 10):
            value = ws.cell(row=row, column=2 + period).value
            assert value == (
                expected_marker if period in item.active_periods else None
            ), (
                item.id,
                period,
                value,
            )
        assert ws.cell(row=row, column=12).value == item.total_text


def test_deliverable_without_milestone_flag_still_gets_diamond() -> None:
    config = load_config(EXAMPLES / "minimal_tech5.json")
    ws = build_workbook(config)["TECH-5"]
    # D-2 is a Deliverable with no is_milestone flag, in period 4.
    row = FIRST_DATA_ROW + 4
    assert ws.cell(row=row, column=1).value == "D-2"
    assert ws.cell(row=row, column=6).value == "♦"
    assert ws.cell(row=row, column=6).font.color.rgb == "00FF0000"


def test_legend_and_notes_present() -> None:
    _, ws = _sample_sheet()
    texts = [
        str(cell.value)
        for row in ws.iter_rows(min_row=FIRST_DATA_ROW)
        for cell in row
        if cell.value is not None
    ]
    joined = "\n".join(texts)
    assert "Legend:" in joined
    assert "Deliverable / Benchmark / Approval" in joined
    assert "Duration of activities shall be indicated" in joined
    assert "Include a legend" in joined


def test_generate_writes_file_with_full_slug(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.chdir(tmp_path)
    output = generate(EXAMPLES / "sample_tech5.json")
    assert output.exists()
    # The original truncated the project name to 20 characters.
    assert "National_Interoperability_Framework" in output.name
    assert output.suffix == ".xlsx"
    wb = openpyxl.load_workbook(output)
    assert wb["TECH-5"].max_row >= FIRST_DATA_ROW + 23


def test_generate_with_explicit_output(tmp_path: Path) -> None:
    output = tmp_path / "nested" / "schedule.xlsx"
    result = generate(EXAMPLES / "minimal_tech5.json", output)
    assert result == output
    assert output.exists()


def test_slugify() -> None:
    assert slugify("My Project: 2026!") == "My_Project_2026"
    assert slugify("   ") == "Output"
    assert len(slugify("x" * 100)) == 60
