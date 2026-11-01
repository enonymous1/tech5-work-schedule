# Changelog

## 2.0.1 — 2026-11-01

Correctness fixes from the 2026-10-31 staff review. No layout, schema,
or API change.

- Fixed guidance note 2 rendering with a double period ("bar chart..").
- Capped `total_duration` at 240: an unbounded duration could build a
  workbook Excel cannot open.
- Default-named outputs no longer overwrite each other within the same
  second: a collision now gets a `-2`, `-3`, ... suffix. Explicit
  `--output` paths still overwrite, as before.
- CLI `generate` reports filesystem errors (permission denied, output
  path is a directory, disk full) as a one-line error with exit code 2
  instead of a traceback.
- Duplicate JSON keys are now rejected instead of silently keeping the
  last value.
- Documented `verify` as a content-only check (it does not check fills,
  fonts, freeze panes, or print setup in this release) and narrowed its
  workbook-load error handling to the failures a bad file produces.
- Exit-code contract documented in `cli.main` (0 success, 1 verify
  mismatch, 2 usage/config/I/O error).

## 2.0.0 — 2026-10-27

Sanitized public export and modernization of the original `mp-tech5` project (formerly a single script pair, `generator.py` + `inspect_totals.py`). The original git history is not carried over; this is a fresh export under the public identity. The original Kosovo configs, rough-draft workbook, and blank template binary are not shipped; the bundled example is fictional data only, with the RFP reference replaced by `{Insert RFP reference}`.

### Added

- Installable `tech5_work_schedule` package with a `tech5-work-schedule` CLI: `generate`, `validate`, and `verify` subcommands, plus `--version`.
- Strict config validation with named errors: unknown/missing keys, duplicate ids, invalid types, and out-of-range periods are rejected instead of silently skipped.
- `verify`, which checks a generated workbook against its config cell by cell (title, subtitle, header, ids, descriptions, markers, Total texts, legend, notes) and reports every mismatch.
- Project name (row 1) and RFP reference + duration (row 2) rendered into the workbook; the workbook title property is set.
- Freeze panes, landscape fit-to-width print setup, and repeating title/header rows.
- Fictional `examples/sample_tech5.json` (23 items, 9 months) and `examples/minimal_tech5.json` (range periods, Sub-Activity, unflagged Deliverable).
- Test suite (config validation, rendering, verify, CLI), ruff + mypy gates, GitHub Actions CI (Python 3.11–3.13), MIT LICENSE, and `.gitattributes` line-ending normalization.

### Fixed

- The old sample config's schema (`project_metadata` / `phases` / `months`) was never read by the generator, which silently produced an empty schedule. That schema is now rejected with guidance toward the supported one.
- Periods of 0 or below overwrote the id/description cells; periods above the total and malformed entries were silently skipped. All periods are now validated against `1..total_duration`, including range bounds and reversed ranges.
- A Deliverable without `is_milestone` rendered as X, contradicting the legend. Deliverables now always render as ♦, as do milestones.
- Workshop bolding by substring match on the description is removed; bolding follows item type only.
- The output filename no longer truncates the project name to 20 characters.
- `inspect_totals.py` silently found nothing on the rough draft it targeted (no Total column, different layout) and exited 0. It is replaced by `verify`, which fails loudly on mismatches.
- The blank one-cell template workbook is dropped; workbooks are built from scratch, so no stale template formatting can leak into outputs.
- Hardcoded config path, CWD-relative template/output paths, and the unpinned `openpyxl` requirement are replaced by CLI arguments and `openpyxl>=3.1` in `pyproject.toml`.
