# Report Refactor Workflow

This project now uses a low-risk, reversible split for the report generator.

## Goal

Reduce the size and change-risk of `generate_report.py` without breaking the public API used by:

- `/Users/rodrigofaria/Documents/VScode/Clinica/Work/server.py`
- local CLI usage of `python3 generate_report.py`

## Current Structure

- `/Users/rodrigofaria/Documents/VScode/Clinica/Work/generate_report.py`
  Compatibility facade and CLI entrypoint.
- `/Users/rodrigofaria/Documents/VScode/Clinica/Work/report_assets.py`
  Frontend HTML template, including CSS and embedded JavaScript.
- `/Users/rodrigofaria/Documents/VScode/Clinica/Work/report_parsers.py`
  Workbook/CSV parsing and shared normalization helpers.
- `/Users/rodrigofaria/Documents/VScode/Clinica/Work/report_builder.py`
  Payload assembly and HTML report builders.
- `/Users/rodrigofaria/Documents/VScode/Clinica/Work/report_flat_time.py`
  Flat-time-specific parsing, dataset loading, and activity-code translation loading.

## Safe Workflow

### Phase 1: Compatibility-first extraction

Done in this refactor:

1. Keep `generate_report.py` import-compatible.
2. Move large, self-contained responsibilities into dedicated modules.
3. Preserve all externally used function names through re-exports in `generate_report.py`.
4. Validate behavior with compile checks and HTML generation smoke tests.

This keeps the deployment surface stable while reducing the blast radius of future edits.

### Phase 2: Focused domain extraction

Done in this refactor:

1. Move flat-time-specific parsing/calculation helpers into `/Users/rodrigofaria/Documents/VScode/Clinica/Work/report_flat_time.py`.
2. Keep existing import paths stable through re-exports from `generate_report.py` and `report_parsers.py`.

Recommended next extractions, only if needed later:

1. Move weekly-report-specific calculations into a dedicated metrics module.
2. Split the frontend template into external static assets if deployment constraints allow it.

These are optional and can be done incrementally.

## Rollback Strategy

This refactor is designed to be reversible.

### Fast rollback

Revert the refactor commit and redeploy.

Because `server.py` still imports `generate_report.py`, rollback does not require changing server wiring.

### Partial rollback

If only one module causes problems:

1. Keep `generate_report.py` as the stable import target.
2. Repoint the affected imports inside `generate_report.py` back to local implementations or a prior module version.
3. Redeploy without touching `server.py`.

## Validation Checklist

Run these after each refactor step:

1. `python3 -m py_compile /Users/rodrigofaria/Documents/VScode/Clinica/Work/generate_report.py /Users/rodrigofaria/Documents/VScode/Clinica/Work/report_assets.py /Users/rodrigofaria/Documents/VScode/Clinica/Work/report_parsers.py /Users/rodrigofaria/Documents/VScode/Clinica/Work/report_builder.py /Users/rodrigofaria/Documents/VScode/Clinica/Work/report_flat_time.py /Users/rodrigofaria/Documents/VScode/Clinica/Work/server.py`
2. Generate an empty dashboard.
3. Generate an XLSX-backed dashboard.
4. Generate a CSV-backed dashboard.
5. Open the result and confirm tabs initialize and charts render.

## Why this is reversible

- External imports stay stable.
- `server.py` does not need to know the internal module layout.
- The largest move was extraction, not redesign of the runtime contract.
- Git rollback remains clean because behavior-preserving moves are isolated from feature work.
