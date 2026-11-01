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

## Current recovery status

The DAILY_INPUT native apply/restore recovery blocker is closed for the tested patient-absence scenario.

Fresh canonical workbook:

`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261101_X.xlsm`

Native scenario:

- patient `ΖΑΛΟΚΩΣΤΑΣ`,
- all-day absence,
- reason/status `διακομισμένος`,
- Apply once,
- Restore once.

Observed green result:

- patient removed correctly from MASTER and THERAPIST_DAILY during active absence,
- patient restored correctly to MASTER and THERAPIST_DAILY,
- `MASTER_SCHEDULE!D8` returned to its internal `PATIENT_PLANNER` formula with no `[1]` external-workbook qualifier,
- `MASTER_SCHEDULE!M8` returned to `παρών`,
- `DAILY_CHANGE_LEDGER` row moved from `ACTIVE` to `RESTORED`,
- `RestoredAt` was populated,
- DAILY_INPUT source row was cleared after restore.

The old W failure was fully explained:

- restore patch #2 targeted `MASTER_SCHEDULE!D8`,
- a formula string had been serialized as a normal value patch,
- VBA compared `cell.Value` against formula text and produced a deterministic read-back mismatch,
- rollback then used cross-workbook Copy and rewrote internal formulas as `[1]PATIENT_PLANNER!...`.

Canonical fix commit:

`70f1564` — `Fix DAILY_INPUT restore formula contract`

Green focused source gate before X:

`17 passed in 4.43s`

Native X proof checkpoint:

`921422f` — `Record green native X restore proof`

## Close-day Y diagnostic outcome

Fresh native close-day candidate:

`Excel/previews/UNIFIED_CLOSE_DAY_HISTORY_20261101_Y.xlsm`

The close-day inspector found no `CURRENT_FINAL` row for `2026-10-30`.

A read-only preparation probe was then run after Excel was fully closed. It reproduced the same deterministic failure:

`DAILY_INPUT contains no active absence/cancellation rows`

No close-day payload was created.

This proves Y did not reach snapshot/history persistence. The blocker was in preparation before any mutation.

Root cause in canonical source:

- `read_daily_input(...)` intentionally supports `allow_empty=True`,
- `prepare_close_day_in_place.py` called it with the default `allow_empty=False`,
- therefore a completely clean operational day with no absence/cancellation rows could not be finalized even if it had zero unresolved sessions.

Bounded fix:

- `4ea05f4` — `Allow clean days in close-day preparation`
- `2b457df` — `Cover clean close-day preparation`

The fix changes only the close-day reader call to `allow_empty=True`. It does not weaken finalization safety: unresolved sessions are still rejected downstream by the existing finalized-day snapshot contract.

Y is frozen as failed diagnostic evidence and must not be reused or repaired in place.

The next source gate must include the close-day preparer/action tests and the native close-day inspector tests. Only after that gate is green should a fresh close-day candidate be built.

## Important interpretation of diagnostic logs

The remaining `%TEMP%` restore / DAILY_INPUT error logs observed after X had timestamps around 18:44 and belong to the earlier failed W run. X restored successfully at approximately 19:06. Do not treat those stale logs as X failures.

## Small post-recovery usability follow-ups

These are not blockers for close-day/history work:

- when a patient is marked `διακομισμένος`, the visible status text may still show strike presentation,
- patient-absence ledger `FromText` currently says only `Κανονικό πρόγραμμα`; a later usability enhancement should show the original therapist and scheduled time, reusing the same scheduling/replacement facts already used in therapist-absence workflows.

Do not reopen the recovery track solely for these cosmetic/history-display enhancements.

## Current application contract

`modCloseDayAction` is already present in `REQUIRED_APP_COMPONENTS` in `Python/rehab_excel/application_contract.py`.

Therefore the older 2026-10-30 close-day checkpoint instruction to promote `modCloseDayAction` into the contract is obsolete.

## Exact next step

Return to the product roadmap:

1. rerun the close-day/history regression gate after the clean-day preparation fix,
2. if green, build a fresh uniquely named close-day workbook,
3. perform one native close-day smoke,
4. verify finalized day snapshots, therapist history and day history events,
5. only after native close-day proof continue to statistics/reporting.

Do not restart DAILY_INPUT verifier debugging unless new native evidence specifically points back there.

## Close-day/history stack already recovered

The recovered close-day implementation includes:

- explicit operator confirmation,
- read-only preparation before mutation,
- in-session persistence,
- `FINALIZED_DAY_SNAPSHOTS`,
- `THERAPIST_DAY_HISTORY`,
- `DAY_HISTORY_EVENTS`,
- duplicate/schema preflight,
- semantic row read-back,
- rollback of partial history writes on failure,
- technical history sheets kept `xlSheetVeryHidden`,
- same-workbook save on success.

Historical close-day checkpoint:

`Documentation/CLOSE_DAY_VBA_RECOVERY_CHECKPOINT_2026-10-30.md`

Use it for implementation history only. Its old next-step section has been superseded by this handoff.

## Native artifact rule

Freeze prior smoke artifacts. Never repair/reuse them.

- W is failed restore evidence only.
- X is green apply/restore evidence only.
- Y is failed close-day preparation evidence only.
- the next native close-day test must use a fresh uniquely named workbook.

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

Continue proactively without waiting for repeated `ok`. Ask for user action only when a genuinely local/native result is required. Keep detailed checkpoint notes for every meaningful outcome so a future conversation never needs to rediscover the same ground.
