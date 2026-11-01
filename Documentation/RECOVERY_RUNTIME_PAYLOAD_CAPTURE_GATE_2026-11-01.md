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

The original file history also shows this flag was present when `prepare_daily_input_in_place.py` was first introduced, so it is not a later regression.

### Nondeterministic patient-centric render ordering

`patient_centric_preview.py` emits patches from:

`sorted(renders.items(), key=lambda item: (item[0][1], item[0][0].casefold()))`

so Python hash/set ordering is not an explanation.

### Recent drift in the patient-centric renderer

The renderer history is old and stable relative to the recovery work. Important commits are from 19–20 September 2026, including `9f0a114`, `e47c6d4`, and `d2f09cd`.

### CHANGE rows causing the patch-number mismatch

The ledger-aware apply path counts only rows whose kind is `PATCH` when it increments `currentPatchIndex`. `CHANGE` rows do not explain native `D8 = patch #2` versus fresh `D8 = patch #11`.

### Patient-centric/business engine failure

The copy-based `Python/tools/apply_daily_input.py` path successfully applied the same U patient absence to a fresh diagnostic workbook.

Observed proof:

- `DAILY INPUT APPLIED PREVIEW OK`
- `Changed sessions: 1`
- affected patient: `planner:4:ΦΘ@2026-10-30 [B7]`
- `Input unchanged: True`
- `VBA preserved: True`
- source SHA256 unchanged before/after.

Therefore the business/planning engine remains functional and must not be rewritten based on the current native rollback symptom.

## Decision

Do not modify replacement logic, THERAPIST_DAILY rendering, MASTER verifier semantics, or start another blind rebuild cycle merely from the contradictory patch number.

First capture the exact payload used by Excel at runtime.

## Diagnostic source change

Added deterministic source migrator:

`Python/tools/apply_daily_input_runtime_capture_source_fix.py`

On an in-place apply failure it changes only the failure path so the exact temp artifacts are copied before normal cleanup:

- `%TEMP%\eka_daily_input_last_payload.tsv`
- `%TEMP%\eka_daily_input_last_response.json`

The existing error log remains:

- `%TEMP%\eka_daily_input_last_error.log`

Business logic, patch construction, patch application, rollback, verifier acceptance rules, and success-path behavior are unchanged.

Regression tests:

`Tests/test_daily_input_runtime_capture_source_fix.py`

The migrator fails closed unless the expected source marker exists exactly once.

## Green source gate

User ran the focused regression suite after applying the source migrator.

Result:

`14 passed in 4.00s`

This is the accepted source gate for the runtime payload capture change.

## Exact next sequence

1. Run `git diff --check`.
2. Commit/push only the canonical `Python/rehab_excel/daily_input_action_vba.py` change.
3. Build one fresh canonical workbook only:
   - `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261101_V.xlsm`
   - using `Python/tools/build_unified_preview.py`
   - no overwrite
   - no post-build patcher.
4. Delete only the old `%TEMP%` diagnostic capture files before native testing.
5. Open V only.
6. Enter the same single patient absence in `DAILY_INPUT`.
7. Press Apply once.
8. If it fails, do not retry V.
9. Immediately capture:
   - `%TEMP%\eka_daily_input_last_error.log`
   - `%TEMP%\eka_daily_input_last_payload.tsv`
   - `%TEMP%\eka_daily_input_last_response.json`
10. Compare the actual runtime payload line-for-line with a fresh manual payload from the same frozen V state before making any verifier change.

## Do not do

- Do not reuse U for another native repair attempt.
- Do not repair-over-repair V after a failure.
- Do not modify replacement/global solver.
- Do not modify patient-centric renderer.
- Do not modify MASTER verifier semantics before exact runtime payload evidence.
- Do not revive N/O/P/Q post-build recovery patching.
- Do not touch the production baseline workbook.

This checkpoint is the authoritative resume point for the next native diagnostic cycle.
