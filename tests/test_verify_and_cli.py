from pathlib import Path

import openpyxl
import pytest

from tech5_work_schedule.cli import main
from tech5_work_schedule.config import load_config
from tech5_work_schedule.generator import FIRST_DATA_ROW, generate
from tech5_work_schedule.verify import verify_files, verify_workbook

EXAMPLES = Path(__file__).resolve().parents[1] / "examples"


def _generated(tmp_path: Path, name: str = "sample_tech5.json") -> Path:
    return generate(EXAMPLES / name, tmp_path / "out.xlsx")


def test_verify_passes_on_generated_workbook(tmp_path: Path) -> None:
    workbook = _generated(tmp_path)
    assert verify_files(workbook, EXAMPLES / "sample_tech5.json") == []


def test_verify_catches_tampered_marker_and_missing_legend(tmp_path: Path) -> None:
    workbook = _generated(tmp_path)
    wb = openpyxl.load_workbook(workbook)
    ws = wb["TECH-5"]
    # Remove a marker that should be present.
    ws.cell(row=FIRST_DATA_ROW, column=3).value = None
    wb.save(workbook)
    mismatches = verify_files(workbook, EXAMPLES / "sample_tech5.json")
    assert any("period 1" in m for m in mismatches)

    # Tamper with the id cell as well: verify reports both problems.
    wb = openpyxl.load_workbook(workbook)
    wb["TECH-5"].cell(row=FIRST_DATA_ROW, column=1).value = "WRONG"
    wb.save(workbook)
    mismatches = verify_files(workbook, EXAMPLES / "sample_tech5.json")
    assert any("expected id" in m for m in mismatches)


def test_verify_missing_workbook_and_bad_config(tmp_path: Path) -> None:
    assert verify_files(tmp_path / "nope.xlsx", EXAMPLES / "sample_tech5.json")
    workbook = _generated(tmp_path)
    assert verify_files(workbook, tmp_path / "missing.json")


def test_verify_workbook_wrong_sheet(tmp_path: Path) -> None:
    wb = openpyxl.Workbook()
    wb.active.title = "Other"
    path = tmp_path / "other.xlsx"
    wb.save(path)
    config = load_config(EXAMPLES / "minimal_tech5.json")
    assert verify_workbook(path, config) == [
        "missing sheet 'TECH-5'; sheets are ['Other']"
    ]


def test_cli_validate_generate_verify_round_trip(
    tmp_path: Path, monkeypatch, capsys
) -> None:
    monkeypatch.chdir(tmp_path)
    assert main(["validate", "--config", str(EXAMPLES / "minimal_tech5.json")]) == 0
    out = capsys.readouterr().out
    assert "Config is valid." in out
    assert "Deliverable: 2" in out

    assert (
        main(
            [
                "generate",
                "--config",
                str(EXAMPLES / "minimal_tech5.json"),
                "--output",
                str(tmp_path / "schedule.xlsx"),
            ]
        )
        == 0
    )
    assert (
        main(
            [
                "verify",
                "--config",
                str(EXAMPLES / "minimal_tech5.json"),
                "--workbook",
                str(tmp_path / "schedule.xlsx"),
            ]
        )
        == 0
    )


def test_cli_generate_with_invalid_config_returns_2(tmp_path: Path, capsys) -> None:
    bad = tmp_path / "bad.json"
    bad.write_text('{"project_meta": {}}', encoding="utf-8")
    assert main(["generate", "--config", str(bad)]) == 2
    assert "error:" in capsys.readouterr().err


def test_cli_verify_mismatch_returns_1(tmp_path: Path, capsys) -> None:
    workbook = _generated(tmp_path, "minimal_tech5.json")
    assert (
        main(
            [
                "verify",
                "--config",
                str(EXAMPLES / "sample_tech5.json"),
                "--workbook",
                str(workbook),
            ]
        )
        == 1
    )
    assert "mismatch:" in capsys.readouterr().err


def test_cli_version(capsys) -> None:
    from tech5_work_schedule import __version__

    with pytest.raises(SystemExit) as excinfo:
        main(["--version"])
    assert excinfo.value.code == 0
    assert __version__ in capsys.readouterr().out
