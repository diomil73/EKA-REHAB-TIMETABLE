# Project Status

Last updated: 2026-11-04

## Entry point

New developers and new contexts start from `Documentation/ONBOARDING.md`. This file tracks milestone state, risks and immediate roadmap only.

Branch: `work/history-statistics-rebuild`.

## Current stage

Active milestone: **statistics/history design checkpoint**.

CREATE is native green. The application-style manager is visible and usable in native Excel. The duplicate outpatient edit bug was reproduced and hardened. The therapist/time collision path and workload refresh were reworked and the latest fresh native smoke was reported by the user as behaving correctly. Workload visibility is present on MASTER and the outpatient edit/conflict/workload milestone is treated as operationally green, subject to broader regression later.

The THERAPIST_DAILY presentation cleanup source gate is now green: **27 passed in 3.09s**. Neutral outpatient text and robotic-orange precedence are covered without changing scheduling truth or workload semantics.

The project now enters the dedicated statistics/history design checkpoint before deep global multi-absence optimizer work. This checkpoint must lock historical facts, weekly therapist workload history, retention and dashboard semantics before implementation begins.

Restore correctness is closed green. Restore performance is deferred.

## Latest confirmed checkpoints

- full regression checkpoint: **699 passed, 19 warnings**
- package relationship normalization + OOXML projection gate: **59 passed in 6.85s**
- latest direct restore backend profile: **16.509s total**
- `build_master_projection`: **0.089s**
- V87-V92 authoritative restore/reconstruction lifecycle: **native green**
- per-row / selected / restore-all in-session restore correctness: **native green**
- previous combined outpatient recurring smoke build: **native green**, trusted source unchanged
- outpatient recurring CREATE: **native green**
- outpatient manager: **native visible and usable**
- duplicate-fix focused gate: **45 passed, 4 warnings in 4.98s**
- collision/workload focused gate: **33 passed, 5 warnings in 4.60s**
- latest native outpatient collision + valid reassignment + workload scenario: **user reported it appears correct**
- THERAPIST_DAILY presentation cleanup gate: **27 passed in 3.09s**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## Closed: restart-safe authoritative daily state

Operational truth is reconstructable from:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics/views`

Replacement facts preserve immutable lifecycle history. Restore correctness, reconstruction and daily statistics parity are native green.

## Closed: Excel ledger authoritative restore correctness

`DAILY_REPLACEMENT_FACTS` is authoritative. `DAILY_CHANGE_LEDGER` is only a projection/audit view.

Primary restore UX keeps the workbook open, runs Python hidden/blocking, protects rollback, imports verified state in-session and avoids visible console / close-reopen behavior.

## Outpatient recurring management + therapist workload UI

Current technical path supports recurring outpatient creation/edit, application-style management, conflict popup/guard, stable same-row identity, daily projection cleanup, workload badges and automatic refresh after outpatient changes.

### Therapist workload strip contract

The workload strip shows one badge per therapist in the form `Όνομα n/6`, where `n` is the number of occupied timeslots in the current operational day. It intentionally counts occupied timeslots, not patient lines, so a legitimate shared slot consumes one capacity unit.

Color thresholds for the normal 6-slot therapist capacity:

- **0-4 / 6: green**,
- **5 / 6: yellow/orange**,
- **6 / 6 or above: red**.

Normal users must not see raw sheet/grid/tab navigation, worksheet row numbers or `outpatient:<row>` IDs.

## Closed source gate: THERAPIST_DAILY presentation consistency

Outpatient text no longer relies on a special font color for identity. The presentation layer keeps outpatient text neutral black while robotic treatment retains orange precedence. The focused gate is green at **27 passed in 3.09s**.

This cleanup is presentation-only and does not alter recurring storage, daily truth, row identity or workload counts.

## H2-H7 history/statistics functional slice — functionally accepted; protection hardening active

The core historical truth and the first user-facing reporting slice are implemented and operator-confirmed on frozen disposable native workbooks:

- **H2 close-day history:** native Z proof; finalized snapshot, therapist/day rows and append-only day event.
- **H3 day amendment:** native AB proof; v1 preserved as SUPERSEDED, v2 CURRENT_FINAL, linked version chain and atomic validation/readback.
- **H4 therapist history:** native AQ proof; daily, weekly, monthly and overall reports, based on CURRENT_FINAL history.
- **H5 patient history:** native AX/AY proofs; first close and v1→v2 amendment with patient-session history.
- **H6 contextual day-save UX:** native AZ/BB proofs; one contextual Save Day action routes first close vs amendment and is exposed on MASTER.
- **H7 combined reporting UI:** native BF manually accepted; integrated RefreshAll + therapist statistics + patient statistics toolbar, report modes and direct previous-screen navigation.

These artifacts are evidence, not production workbooks. Preserve them frozen. The current H7 protection follow-up is NOT yet accepted: BG and BH are failed/frozen; BH protected the therapist report but exposed that the patient-report installer left PATIENT_HISTORY_REPORT unprotected. Source corrections are committed, but the combined **33-test** gate has not yet been operator-run. The next candidate is a fresh **BI from trusted AZ**, only after all 33 tests pass and Excel is closed. Then perform a complete read-only audit before any manual clicks.

See `Documentation/SESSION_RECOVERY_H7_PROTECTION_BG_FAILURE_2026-11-04.md` for the exact gate and artifact rules.

## Active: final worksheet edit-protection hardening

The accepted BF audit found 23 worksheets: 18 with `ProtectContents=False` and five protected. Four technical history sheets are already protected and veryHidden; PATIENT_HISTORY_REPORT is protected in BF; THERAPIST_HISTORY_REPORT was not. The targeted therapist and patient report installation gaps are being corrected first.

No broad worksheet protection has been implemented. The full embedded VBA write-surface inventory has not yet been supplied/reconciled, and the intended human-edit zones are not yet established for every sheet. Do not blanket-protect or lock workbook structure until dynamic sheet creation, sanctioned form writes, external COM writers, macros, rollback and error paths are classified.

The source-derived partial map is `Documentation/WORKSHEET_PROTECTION_WRITE_PATH_MAP_2026-11-04.md`. Its remaining read-only VBA inventory command and the exact 33-test / BI gate are in the H7 recovery packet.

## Open: replacement quality under multiple therapist absences

Replacement lifecycle correctness is green, but global assignment quality remains open.

Required global behavior:

- enter all known therapist absences first,
- solve all displaced patients together,
- maximize placements / minimize unresolved orphans,
- protect unique viable slots from flexible patients,
- prefer original time when higher-priority constraints permit,
- preserve ΦΘ/reclining proximity where feasible,
- avoid significant provider-load imbalance,
- respect patient absences and other treatments,
- add persistent pool/robotic operating-hour configuration,
- treat 13:00 as a weak soft preference to avoid,
- permit double-booking/over-capacity only through explicit last-resort user approval.

Canonical scarcity example: A can use 09:00/10:00/11:00, B only 09:00, so protect 09:00 for B when a complete feasible solution exists.

This must be addressed before replacements are considered complete.

## Deferred: restore performance

Current end-user restore duration in the low-to-mid 20-second range is accepted for this stage. Do not resume performance work unless there is a regression or final-polish phase.

## Immediate roadmap

1. Hold the statistics/history design session now and lock the historical fact model before coding it.
2. Define weekly therapist workload history, including both occupied timeslots and active-session counts.
3. Define finalized-day snapshot/event semantics so later schedule edits do not rewrite history.
4. Define the first user-facing statistics/dashboard views and comparison periods.
5. Implement the smallest persistent historical slice with tests before adding broader reports.
6. Resume broader scheduling/UI work after the history contract is stable.
7. Implement global multi-absence optimization and persistent special-service operating hours before replacements are considered complete.
8. Revisit restore performance near completion.

## Locked daily rules

1. recurring/base schedule,
2. cancellations/no-shows release slots,
3. therapist absences,
4. identify displaced/orphan sessions,
5. apply surviving ACTIVE replacements,
6. unresolved orphans resolve through valid placement or treatment suspension,
7. only then expose capacity to new patients.

Capacity is timeslot-based. Existing displaced patients precede new patients. MASTER and THERAPIST_DAILY must agree.

## Safety constraints

- baseline `.xlsm` is immutable,
- preview/smoke workflows use fresh copies,
- preserve VBA and source hash,
- keep write zones narrow,
- bulk restore validates the entire requested set before mutation,
- no production/staging `.xlsm` save with openpyxl,
- no manual Save in independent smoke scenarios,
- Windows tests use `python -m pytest`.

## Documentation discipline

`Documentation/ONBOARDING.md` is the single front door. Specialist files keep their own source-of-truth roles. Meaningful changes must update the relevant specialist document in the same cycle, and ONBOARDING when the project-wide mental model changes.


## Latest milestone update — 2026-11-04

- H7 patient/therapist history/statistics: **source GREEN and native manual acceptance confirmed** on disposable BF.
- Integrated toolbar includes retained RefreshAll, therapist statistics, patient statistics, with direct return navigation from both reports.
- BF independent read-only integrated audit passed; BF remains frozen smoke evidence, not production.
- Next active milestone: final worksheet edit-protection hardening, beginning with a read-only protection inventory. No new sheet protection has yet been applied.
- Exact status and resume path: `Documentation/CURRENT_HANDOFF.md` and `Documentation/CHECKPOINT_H7_BF_MANUAL_ACCEPTANCE_AND_PROTECTION_NEXT_2026-11-04.md`.


## Protection milestone update — 2026-11-04 (BG failed; retry gated)

- BF remains accepted, frozen native smoke evidence. AZ remains trusted source. Neither was modified.
- Full BF protection audit completed read-only: 23 sheets; 18 unprotected, five protected; workbook unchanged.
- Targeted therapist report protection was added because the blank visible report sheet was unprotected in BF.
- BG integrated preparation failed at therapist report installation with a drawing-object protection verification failure. BG is frozen failed evidence and must not be repaired, reused or overwritten.
- A stale source-inspection test was corrected to inspect the persisted-audit helper. COM protection call was changed to explicit positional arguments, and failure output now reports actual `ProtectContents` / `ProtectDrawingObjects` state. Installer also verifies each report button binding and all four technical history sheets remain veryHidden.
- Current focused suite: 15 tests, not yet rerun after the latest corrections. Next gate is documented in `Documentation/CURRENT_HANDOFF.md`: pull, run `Tests/test_therapist_history_ui.py`, expect 15 passed. Only then create a new BH from AZ. Freeze BH on any error.
- No broad protection has been applied to the other 17 unprotected sheets. Do not expand protection until this targeted report path is native-green and all manual edit zones are classified.


Recovery entry point for the failed BG / pending BH therapist-report protection work:
`Documentation/SESSION_RECOVERY_H7_PROTECTION_BG_FAILURE_2026-11-04.md`.
It records frozen artifacts, commit trail, the exact 15-test gate, next BH command, and cross-session read order. BF remains accepted/frozen, AZ trusted/frozen, BG failed/frozen. Do not create BH until all 15 focused tests pass.


## 2026-11-04 — therapist report protection test gate GREEN

Operator confirmed the latest focused suite: `15 passed in 4.11s`. The native BH smoke is now authorized from trusted AZ. BF remains frozen accepted evidence; BG remains frozen failed evidence; AZ remains unchanged. Follow the exact safety-checked commands and audit expectations in `Documentation/SESSION_RECOVERY_H7_PROTECTION_BG_FAILURE_2026-11-04.md`. Do not proceed to manual use or broader worksheet protection until BH's read-only audit confirms the narrow protection delta.

## 2026-11-04 — BH failed; patient-report installer regression isolated

The new BH native workbook was created from trusted AZ after the therapist-only 15-test gate passed. Its read-only protection audit confirmed therapist report protection (`ProtectContents=True`, `ProtectDrawingObjects=False`) and all four technical history sheets protected/veryHidden, but exposed an unintended regression: `PATIENT_HISTORY_REPORT` was unprotected. The accepted BF baseline had this report protected.

Root cause confirmed from source: `install_patient_report_components()` calls `report.Unprotect()` and clears the sheet, but does not re-protect the blank report before returning. The patient installer audit did not check this persisted state. The report VBA renderer only restores protection after a render, so previous manual use in BF masked the gap.

Source changes now committed:
- `81a8615`: two regression tests for patient report installer-time and persisted protection.
- `0fd7435`: protect blank patient report with `Protect("", False, True, True)`; verify contents protected and drawing objects unprotected.
- `61ebf11`: persisted patient report protection check in patient installer audit.

Operator test result is pending. BH is frozen failed evidence; do not repair/reuse. BF accepted/frozen; AZ trusted/frozen; BG failed/frozen. Next gate is 33 focused tests across therapist UI, patient UI, integrated toolbar/order and previous-screen navigation. Only after green create fresh BI from AZ and run read-only audit. No manual testing or broader sheet protection before the audit passes.

