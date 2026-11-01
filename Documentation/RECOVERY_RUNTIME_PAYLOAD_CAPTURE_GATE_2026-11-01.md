# Recovery runtime payload capture gate — 2026-11-01

## Why this gate exists

Native U failure evidence and fresh manual payload evidence disagree on patch ordering:

- native U log at 2026-11-01 12:37:33 reports `patch_index=2`, `MASTER_SCHEDULE!D8`,
- a fresh payload prepared from the same U workbook reports `MASTER_SCHEDULE!D8` as raw `PATCH #11` and `MASTER_SCHEDULE!M8` as `PATCH #12`.

This discrepancy must be resolved before changing verifier semantics again.

## What has been ruled out

### Wrong workbook VBA / wrong script path

`inspect_u_runtime_provenance.py` proved:

- U contains the character-strike helper,
- U contains the bounded tri-state non-active overlay branch,
- U contains the whole-cell-first verifier helper,
- `ResolveDailyInputPreparationScript()` resolves to the repository `Python/tools/prepare_daily_input_in_place.py`,
- resolved script SHA256 equals repository script SHA256.

### Late addition of `rebuild_multi_member_groups=True`

`git blame` proves the flag exists from the initial in-place preparer commit:

`17d55687 (2026-10-30 23:29:40 +0200) rebuild_multi_member_groups=True`

### Nondeterministic patient-centric render ordering

`patient_centric_preview.py` emits patches from:

`sorted(renders.items(), key=lambda item: (item[0][1], item[0][0].casefold()))`

so Python hash/set ordering is not an explanation.

### Recent drift in the patient-centric renderer

The renderer history is old and stable relative to the recovery work. Important commits are from 19–20 September 2026, including `9f0a114`, `e47c6d4`, and `d2f09cd`.

## Decision

Do not modify replacement logic, THERAPIST_DAILY rendering, MASTER verifier semantics, or rebuild another native artifact merely from the contradictory patch number.

First capture the exact payload used by Excel at runtime.

## Diagnostic source change

Added deterministic source migrator:

`Python/tools/apply_daily_input_runtime_capture_source_fix.py`

On an in-place apply failure it changes only the failure path so the exact temp artifacts are copied before normal cleanup:

- `%TEMP%\eka_daily_input_last_payload.tsv`
- `%TEMP%\eka_daily_input_last_response.json`

Business logic, patch construction, patch application, rollback, and success-path behavior are unchanged.

Regression tests:

`Tests/test_daily_input_runtime_capture_source_fix.py`

The migrator fails closed unless the expected source marker exists exactly once.

## Required local gate

1. Pull branch.
2. Run the source migrator once.
3. Run focused regression tests.
4. Commit only the resulting canonical source change after green tests.
5. Only then build one fresh diagnostic workbook.
6. Run one patient absence Apply.
7. If it fails, inspect `%TEMP%\eka_daily_input_last_payload.tsv` and compare its numbered `PATCH` lines directly with the error log.

Do not reuse U or any prior failed workbook for the diagnostic native run.
