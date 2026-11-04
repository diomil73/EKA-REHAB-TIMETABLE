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

## DAILY_INPUT recovery status

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

Canonical fix commit:

`70f1564` — `Fix DAILY_INPUT restore formula contract`

Native X proof checkpoint:

`921422f` — `Record green native X restore proof`

## Close-day durable history status — H2 closed

The earlier Y candidate failed before persistence because clean days were rejected by DAILY_INPUT preparation. Y is frozen failed evidence.

Bounded fix:

- `4ea05f4` — `Allow clean days in close-day preparation`
- `2b457df` — `Cover clean close-day preparation`

Fresh green close-day/history proof:

`Excel/previews/UNIFIED_CLOSE_DAY_HISTORY_20261101_Z.xlsm`

Operator-confirmed native Z flow completed successfully:

- focused close-day/history regression gate green,
- fresh Z build green,
- read-only preflight green,
- native `CloseOperationalDay` executed once successfully,
- post-close inspector green.

Exact inspector field values were not pasted, so do not invent them.

Checkpoint:

`eed0eb0` — `Record green native Z close-day proof`

H2 durable close-day persistence is therefore closed/operator-confirmed.

## Closed-day correction/version chain — H3 backend green

The H3 backend now has operator-confirmed native proof.

Focused source gate:

`35 passed in 3.23s`

The implemented H3 stack includes:

- pure-core amendment/version-chain contract,
- preservation of previous finalized snapshot,
- explicit `SUPERSEDED` status for the prior version,
- replacement `CURRENT_FINAL` snapshot with exactly `version + 1`,
- `supersedes_snapshot_id` linkage,
- append-only `DAY_AMENDMENT_STARTED` and `DAY_AMENDMENT_FINALIZED` events,
- atomic native Excel persistence in one COM session,
- no `Save()` until every preflight and semantic read-back passes,
- read-only amendment preparation,
- source SHA-256 binding between preparation and apply,
- post-save semantic verification,
- independent amendment outcome inspector.

Fresh green native artifact:

`Excel/previews/UNIFIED_DAY_AMENDMENT_20261101_AB.xlsm`

Observed inspector output:

```text
DAY AMENDMENT OUTCOME OK
Date: 2026-10-30
Version: 1 -> 2
Current snapshot: finalized:2026-10-30:v2
Therapist/day history rows: 16
```

This proves:

- version 1 is preserved as superseded history,
- version 2 is the new current final snapshot,
- the version chain is contiguous,
- the current amended snapshot has zero unresolved sessions by contract,
- therapist/day history exists for v2,
- amendment events are present and correctly linked under the inspector contract.

Checkpoint:

`19fd49c` — `Record green native AB day amendment proof`

AB is frozen green evidence and must not be reused or repaired in place.

AA is not authoritative proof. A safety stop showed AA already existed, after which a copy command was still issued. Freeze AA as uncertain evidence and never reuse it.

## App-wide navigation convention

`Save & Exit` is an application-shell action, not a MASTER-only feature.

Canonical behavior comes from `modMasterToolbar.ExitApplication` and appears on:

- `MASTER_SCHEDULE`,
- `DAILY_INPUT`,
- `THERAPIST_DAILY`,
- `REPLACEMENTS`.

It saves the workbook, restores the normal Excel interface, and exits/closes the application workbook.

`Κλείσιμο ημέρας` and future `Διόρθωση κλεισμένης ημέρας` are business operations. They must remain distinct from `Save & Exit` and must not duplicate or silently replace it.

## Important interpretation of diagnostic logs

The remaining `%TEMP%` restore / DAILY_INPUT error logs observed after X had timestamps around 18:44 and belong to the earlier failed W run. X restored successfully at approximately 19:06. Do not treat those stale logs as X failures.

## Small post-recovery usability follow-ups

These are not blockers for history/statistics work:

- when a patient is marked `διακομισμένος`, the visible status text may still show strike presentation,
- patient-absence ledger `FromText` currently says only `Κανονικό πρόγραμμα`; a later usability enhancement should show the original therapist and scheduled time, reusing the same scheduling/replacement facts already used in therapist-absence workflows.

Do not reopen the recovery track solely for these cosmetic/history-display enhancements.

## Current application contract

`modCloseDayAction` is already present in `REQUIRED_APP_COMPONENTS` in `Python/rehab_excel/application_contract.py`.

Therefore the older 2026-10-30 close-day checkpoint instruction to promote `modCloseDayAction` into the contract is obsolete.

## Exact next step

Continue H3 with the explicit user-facing closed-day correction workflow.

The next implementation slice must:

1. expose a separate `Διόρθωση κλεισμένης ημέρας` action in the application shell,
2. keep correction separate from ordinary DAILY_INPUT editing and from `Save & Exit`,
3. require an explicit amendment reason/summary before persistence,
4. prepare the amendment read-only first,
5. apply only the hash-bound prepared payload,
6. preserve the existing v1→v2 atomic/version-chain guarantees,
7. use a fresh uniquely named workbook for every native UI smoke,
8. only after native UI proof move on to H4 historical reporting/statistics.

Do not restart DAILY_INPUT or close-day recovery debugging unless new evidence specifically points there.

## Roadmap after H3 UI integration

### H4 — first historical report slice

Report at least:

- therapist,
- day,
- occupied timeslots,
- active-session count,
- final operational outcome counts,
- actual workload,
- planned workload when available.

Then add weekly therapist workload history/trend.

### H5 — patient historical statistics

Add:

- scheduled/delivered,
- missed/no-show,
- department cancellation / specific-session cancellation,
- patient absence days/reasons,
- replacement/suspension outcomes,
- later 15-day progress trend.

### H6 — therapist profile/statistics

Add:

- current workload,
- historical workload/trend,
- absence days/reasons,
- replacements/history,
- scheduled vs delivered.

## Native artifact rule

Freeze prior smoke artifacts. Never repair/reuse them.

- W: failed restore evidence only.
- X: green apply/restore evidence only.
- Y: failed close-day preparation evidence only.
- Z: green close-day/history proof only.
- AA: uncertain amendment candidate, never reuse.
- AB: green H3 amendment/version-chain proof only.

Every new native UI test must use a fresh uniquely named workbook.

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
- finalized day must always have unresolved/orphan sessions = 0,
- no broad THERAPIST_DAILY clearing,
- no blind alphabetic rebuild loops,
- Python tests before native proof,
- native proof before integration,
- no silent destructive rewrite of historical snapshots.

## User working preference

Continue proactively without waiting for repeated `ok`. Ask for user action only when a genuinely local/native result is required. Keep detailed checkpoint notes for every meaningful outcome so a future conversation never needs to rediscover the same ground.


## Latest checkpoint — 2026-11-04: H7 accepted, protection audit next

H7 patient/therapist statistics and integrated toolbar are now **user-confirmed functional** on disposable BF:
`Excel/previews/UNIFIED_PATIENT_HISTORY_H7_20261104_BF.xlsm`.

Evidence:
- 26 focused source tests passed in 4.87s.
- BF independent read-only integrated audit passed.
- User manually confirmed all functions worked correctly: toolbar placement, RefreshAll, both statistics reports, and direct return to THERAPIST_DAILY.
- RefreshAll is retained; it invokes six operational procedures.
- Both history report sheets are visible; technical history sheets remain veryHidden.

Do not patch or reuse BF; it is frozen green smoke evidence, not production.

**Current active task: final worksheet edit-protection hardening.**
Read:
- `Documentation/CHECKPOINT_H7_BF_MANUAL_ACCEPTANCE_AND_PROTECTION_NEXT_2026-11-04.md`
- `Documentation/BACKLOG_FINAL_WORKSHEET_PROTECTION_2026-11-03.md`

Next exact command after pull and with Excel closed:
```powershell
python -m pytest Tests/test_workbook_protection_audit.py -q
python .\Python\tools\audit_workbook_protection.py `
  ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H7_20261104_BF.xlsm"
```
The new audit is read-only. Share full output. Do not implement blanket protection until each sheet's intended edit zones and macro/form write zones are mapped. Protecting sheets is the next milestone, but no Protect/Unprotect changes have yet been made.

## Protection audit follow-up — 2026-11-04

The first five-test run had 4 passed, 1 failed because the test expected an explicit `"UsedRange.Locked"` source literal. The audit already queried the value; diagnostic output is now a dedicated `UsedRange.Locked:` label. Protection allowance reporting now runs for every worksheet. Fix commit: `1d65d90`.

Operator also issued the read-only audit command but did not paste its output. Do not infer results. Pull `1d65d90`, rerun `python -m pytest Tests/test_workbook_protection_audit.py -q`, then rerun the read-only audit on frozen BF and provide the full output. No sheet protection changes yet.
