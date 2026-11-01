# Changelog

## 2.1.0 — 2026-11-01

Structural release from the 2026-10-31 staff review. Generated workbooks
are byte-for-byte equivalent in layout to 2.0.1 (same default theme);
the changes are internal, plus `verify` now also checks styling.

- New `markers` module: `marker_for()` is the single source of truth
  for the legend's marker rule; renderer, verifier, and tests import it.
- New `layout` module: a frozen `Layout` dataclass computes Total
  column, data rows, freeze panes, and print titles once per workbook.
- New `theme` module: a frozen `Theme` dataclass holds colours, fonts,
  and widths; `build_workbook(config, theme=...)` renders with it and
  `verify_workbook(path, config, theme=...)` checks against it. The
  default theme reproduces the 2.0.0/2.0.1 look exactly.
- `generator.py` split into `render.py` (workbook rendering) and
  `io.py` (filenames and file saving); `generator` remains as a
  backwards-compatible re-export shim.
- `ScheduleItem.type` is now a `StrEnum` (`ItemType`); comparisons with
  plain strings still work.
- Optional `schema_version` key (absent means 1; any other value is
  rejected) so a future schema change can keep accepting old configs.
- Long descriptions now wrap and grow their row height instead of
  being clipped at the column-width cap; guidance-note row heights are
  estimated from the final column widths instead of fixed at 45/25/20.
- `verify` now checks key styling against the theme (header/phase/
  deliverable fills, milestone marker colour and boldness, font name,
  freeze panes, page orientation, paper size, fit-to-width, print
  titles). A style-stripped workbook now fails verification.
- `slugify` is Unicode-aware: accented letters survive, only path
  separators, control characters, and Windows-reserved characters are
  removed.
- Public API exported from the package root (`Config`, `ItemType`,
  `Layout`, `Theme`, `build_workbook`, `generate`, `load_config`,
  `marker_for`, `verify_files`, `verify_workbook`, and more).
- CI now also runs on Windows (Python 3.12) and checks that the package
  builds (`python -m build`, `twine check`). The README was replaced
  with the staff-review proposal, updated for 2.1.0.

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
