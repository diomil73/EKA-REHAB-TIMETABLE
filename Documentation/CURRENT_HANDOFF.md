# Current Handoff

Last updated: 2026-11-01

## Mandatory read order for any new conversation or developer

1. `Documentation/ONBOARDING.md`
2. `Documentation/PROJECT_CONTINUITY_MASTER_2026-11-01.md`
3. this file
4. the newest dated recovery checkpoint from 2026-11-01

The master continuity document is authoritative for product goals, safety rules, architecture, recovery history, failed approaches, proven behavior, and exact resume procedure.

## Current branch

`work/history-statistics-rebuild`

## Current product compass

The system exists to:

- simplify patient entry, discharge, absence, return, treatment changes and therapist changes,
- provide a complete daily operational picture of the rehabilitation center,
- preserve patient and therapist history/statistics,
- track patient progress over time,
- link the timetable, patient record/folder and patient progress/history view.

Do not let a local technical workaround undermine those goals.

## Most important current forensic conclusion

The patient-centric / DAILY_INPUT planning engine is still functional.

On 2026-11-01 the user ran the old copy-based apply path against the current U workbook:

- input: `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261101_U.xlsm`
- output: `Excel/previews/DIAGNOSTIC_OLD_ENGINE_FROM_U_20261101.xlsm`
- tool: `Python/tools/apply_daily_input.py`

Result:

- `DAILY INPUT APPLIED PREVIEW OK`
- patient rows read: 1
- operational absences: 1
- changed sessions: 1
- `patient_absent: planner:4:ΦΘ@2026-10-30 [B7]`
- input unchanged: True
- VBA preserved: True
- input hash unchanged: True

Checkpoint: `f6e77bd89f3dd2ed3b0ca79b275d7d265510f382`

Therefore the current failure is NOT evidence that the replacement/patient-centric engine is broken.

## Active blocker

The newer in-place VBA transaction applies the prepared payload, then fails semantic read-back at the first MASTER treatment presentation patch and rolls the transaction back.

Observed error log pattern:

- phase: `verify patch`
- patch_index: `2`
- sheet: `MASTER_SCHEDULE`
- cell: D8 or equivalent affected treatment cell
- error: `DAILY_INPUT read-back mismatch`

Restore then fails secondarily because it uses the same validated patch engine.

This rollback hides all downstream effects, which explains why the user sees no durable THERAPIST_DAILY source/destination changes, ledger progression or restore behavior.

## Exact next step

Do not build another V/W artifact yet.

Do not modify the replacement solver.

Do not modify patient-centric THERAPIST_DAILY rendering.

Instrument the VBA verifier so the failing MASTER patch logs each individual verification predicate before rollback:

- expected/actual value,
- source tag,
- strike state,
- font color,
- text-run state,
- source-tag presentation result,
- generic strike verification result,
- exact branch/Exit Function responsible for failure.

The next source change must be based on that exact predicate, not another blind verifier guess.

## Recent source gate

Bounded MASTER tri-state/read-back source tests reached:

`63 passed, 2 warnings in 3.88s`

The warnings are known Python 3.14 ZipFile cleanup warnings and are not current blockers.

## Native artifact rule

Freeze all prior smoke artifacts. Never repair/reuse them.

U is evidence only.

A new native artifact is created only after the next diagnostic/source fix passes its focused Python regression gate.

## Retired approach

Do not revive the post-build recovery patch/wrapper path:

- `Python/tools/patch_native_restore_recovery.py`
- `Python/tools/build_unified_preview_native_recovery.py`

Normal builds use only:

`Python/tools/build_unified_preview.py`

## Safety essentials

- production baseline `Rehab_Center_System_v27_1.xlsm` is immutable,
- unique fresh preview for every native test,
- preserve source hash and VBA,
- no broad THERAPIST_DAILY clearing,
- no blind alphabetic rebuild loops,
- Python tests before native proof,
- native proof before integration.

## User working preference

Continue proactively without waiting for repeated “ok”. Ask for user action only when a genuinely local/native result is required. Keep detailed checkpoint notes for every meaningful outcome so a future conversation never needs to rediscover the same ground.
