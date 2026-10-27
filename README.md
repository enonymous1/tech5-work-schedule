# tech5-work-schedule

Generate a World Bank **TECH-5 — Work Schedule and Planning for Deliverables** workbook from a structured JSON config: one config in, one formatted Excel bar chart out. Phases, activities, and deliverables are laid out across the schedule's time units, with a legend and the form's guidance notes, ready to drop into a technical proposal.

This tool is unaffiliated with the World Bank. TECH-5 is one of the standard forms in Section 3 (Technical Proposal) of the World Bank's standard procurement documents. The bundled examples hold fictional data only.

Formerly `mp-tech5`; renamed for publication. See [CHANGELOG.md](CHANGELOG.md).

## Quick start

```bash
pip install -e .
tech5-work-schedule validate --config examples/sample_tech5.json
tech5-work-schedule generate --config examples/sample_tech5.json --output schedule.xlsx
tech5-work-schedule verify --config examples/sample_tech5.json --workbook schedule.xlsx
```

With no `--output`, `generate` writes to `output/TECH5_<project>_<timestamp>.xlsx` under the current directory.

Requires Python 3.11+ and openpyxl. No other runtime dependencies.

## Config format

```json
{
  "project_meta": {
    "project_name": "National Interoperability Framework - Example Schedule",
    "rfp_reference": "{Insert RFP reference}",
    "time_unit_label": "Months",
    "total_duration": 9
  },
  "schedule_items": [
    {
      "id": "A-1",
      "description": "PHASE 1: INCEPTION & MOBILIZATION",
      "type": "Phase",
      "active_periods": [1],
      "total_text": "2 Weeks"
    },
    {
      "id": "D-1",
      "description": "DELIVERABLE 1: INCEPTION REPORT (Due Week 2)",
      "type": "Deliverable",
      "active_periods": [1],
      "is_milestone": true,
      "total_text": "Week 2"
    }
  ]
}
```

- `project_meta.project_name` and `total_duration` are required. `rfp_reference` and `time_unit_label` (default `Months`) are optional. Both are rendered into the workbook: row 1 is the project name, row 2 is the RFP reference and duration.
- `type` is one of `Phase`, `Activity`, `Deliverable`, or `Sub-Activity`.
- `active_periods` holds period numbers, or `"start-end"` range strings such as `"2-3"`. Every period must be within `1..total_duration`.
- `is_milestone` (default `false`) and `total_text` (free text for the Total column, e.g. `"Week 10"`) are optional.
- Markers follow the legend: █ for phases, X for activities, and a red ♦ for deliverables and milestones.

Validation is strict and fail-loud. Unknown keys, duplicate ids, missing keys, and out-of-range periods are errors with the config path named, instead of being silently skipped or overwriting other cells, as the original did. `validate` checks a config without generating anything; `verify` checks a generated workbook against its config, cell by cell.

## Examples

- `examples/sample_tech5.json` — a 23-item, 9-month fictional schedule (phases, activities, deliverables, workshops, total texts). This is the example to start from.
- `examples/minimal_tech5.json` — a 5-item, 4-month config exercising range periods, a Sub-Activity, and a Deliverable with no `is_milestone` flag.

The original project's Kosovo configs and rough-draft workbook were real proposal material and are not shipped here. The sample above keeps their structure and durations with fictional, generalized wording, and the RFP reference replaced by a placeholder.

## What changed in 2.0.0

See [CHANGELOG.md](CHANGELOG.md) for the full list. In short: installable package and CLI instead of hardcoded scripts, one validated schema, period validation, the project name and RFP reference rendered in the sheet, print setup and freeze panes, `verify` replacing the old inspector (which silently found nothing on its target file), and the blank template binary dropped — the workbook is built from scratch.

## Development

```bash
pip install -e ".[dev]"
ruff check src tests && ruff format --check src tests
mypy src
pytest -q
```

## License

MIT. See [LICENSE](LICENSE).
