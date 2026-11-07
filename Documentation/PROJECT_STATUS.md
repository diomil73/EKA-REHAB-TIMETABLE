# Project Status

Last updated: 2026-11-04

## Entry point

New developers and new contexts start from `Documentation/ONBOARDING.md`. This file tracks milestone state, risks and immediate roadmap only.

Branch: `work/history-statistics-rebuild`.

## Current stage

Active milestone: **H7 report-protection correction, followed by final worksheet edit-protection hardening**.

The operational scheduling/reconstruction and H2-H7 history/statistics functional slices have substantial native green evidence. H7's integrated toolbar, RefreshAll, patient and therapist report surfaces, report modes and previous-screen navigation were manually accepted on frozen disposable BF. The remaining immediate issue is not report functionality: the later BH protection audit exposed a patient-report installer protection regression. The source correction is committed, but the 33-test gate and fresh BI native proof are still pending.

CREATE is native green. The application-style outpatient manager, recurring edit/conflict handling and therapist workload visibility are treated as operationally green based on the documented native evidence. The global multi-absence assignment optimizer remains open.

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

1. Close the H7 report-protection regression: run the 33 focused tests, create fresh BI from AZ only if green, perform the persisted read-only audit, then manually verify both reports and toolbar.
2. Finish the read-only workbook-embedded VBA write-surface audit and reconcile it with the tracked Python/VBA map.
3. Define per-sheet protection policy and approved human input ranges; preserve controlled macro/backend writes, dynamic sheet creation, save/rollback and report buttons.
4. Implement and natively prove final worksheet protection on fresh disposable workbooks. This is a mandatory production-release gate.
5. Implement the global multi-therapist-absence replacement optimizer and persistent pool/robotic operating-hour configuration before calling replacement quality complete.
6. Reconcile the broader product roadmap against current code: patient 15-day progress/evaluation, patient/therapist search and profiles, doctor assignment/history, archive/reactivation flows, quick lists, intranet links, and portable pilot/exhibition packaging.
7. Complete integrated regression, release packaging, backup/versioning and second-machine pilot validation.

The H2-H7 history/statistics core is not waiting to be designed; it is implemented and functionally accepted. The active work is the H7 protection correction and then broader worksheet hardening. Restore performance remains deferred unless it regresses or final release polish requires it.

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

## Authoritative latest update — H7 33-test gate GREEN; BI next (2026-11-04)

Operator confirmed the combined source gate: **33 passed in 4.13s**. Current test-file breakdown verified from source: therapist UI 15, patient UI 7, integrated toolbar/patient-order 8, previous-screen navigation 3.

The next authorized action is a fresh BI native smoke from trusted AZ, followed immediately by the complete read-only workbook protection audit. BI must not be manually used unless that audit passes. AZ and BF remain frozen; BG and BH remain frozen failed evidence. Exact commands and acceptance criteria are in `Documentation/SESSION_RECOVERY_H7_PROTECTION_BG_FAILURE_2026-11-04.md` and the final section of `Documentation/CURRENT_HANDOFF.md`.

This supersedes the older pending-test statement above. No broad worksheet protection is authorized yet.


## Latest operator evidence — BI integrated install and protection audit GREEN (2026-11-07)

The operator supplied the actual PowerShell output for fresh `Excel/previews/UNIFIED_PATIENT_HISTORY_H7_20261104_BI.xlsm`, built from trusted `UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm` after the `33 passed in 4.13s` gate. The output is chronological: an older BH audit is followed by the new BI build and audit; do not confuse the BH failure with the current BI result.

BI preparation reported:
- `H7 INTEGRATED COPY VERIFIED`
- `UNIFIED DAY SAVE UI INSTALL OK`; persisted VBA reopen verification OK
- `THERAPIST HISTORY UI INSTALL OK`
- `H7 PATIENT HISTORY UI INSTALL OK`; retained/styled RefreshAll and aligned two statistics buttons
- `Persisted VBA reopen verification: OK`
- `H7 INTEGRATED UI INSTALL OK`
- `Source unchanged; fresh destination installed`

The complete separate BI read-only audit reported:
- `ReadOnly: True`; `Worksheet count: 23`; `ProtectStructure: False`; `ProtectWindows: False`
- `PATIENT_HISTORY_REPORT`: Visible=-1, ProtectContents=True, ProtectDrawingObjects=False, ProtectScenarios=True
- `THERAPIST_HISTORY_REPORT`: Visible=-1, ProtectContents=True, ProtectDrawingObjects=False, ProtectScenarios=True
- `FINALIZED_DAY_SNAPSHOTS`, `THERAPIST_DAY_HISTORY`, `PATIENT_SESSION_HISTORY`, `DAY_HISTORY_EVENTS`: Visible=2 (veryHidden), ProtectContents=True, ProtectDrawingObjects=True
- The remaining 17 sheets preserve their expected unprotected baseline; no broad protection was attempted
- `Workbook unchanged: True`
- `WORKBOOK PROTECTION AUDIT COMPLETE`

**Disposition:** BI source gate + native install/persisted-state + read-only worksheet protection audit **GREEN**. The BH patient report regression is resolved at persisted worksheet-state level. This does NOT yet prove that the protected report buttons work when actually clicked, or that a report-mode refresh under protection succeeds. Native manual acceptance is the next and only current H7 step; BI is the accepted audit candidate, not yet overall H7 hardening completion.

### Next manual acceptance, BI only

Open the already-built BI in Excel, macros enabled. Verify the existing application/THERAPIST_DAILY toolbar and RefreshAll. From THERAPIST_DAILY open therapist statistics and check daily, weekly, monthly, overall and the previous-screen button back to THERAPIST_DAILY. Repeat for patient statistics. Confirm both protected reports remain read-only to manual cell edits, buttons remain clickable and four technical history sheets remain hidden. Use the normal report UI only; no production data edit, day-close, amendment or unrelated workflow in this smoke.

If any button/mode fails, stop immediately, record exact visible error, affected report/mode and workbook state; freeze BI as failed evidence, do not patch/retry on BI. If all pass, record the operator's actual manual results and freeze BI as green accepted evidence. Move next to the broad worksheet protection write-path classification; do not lock other 17 sheets until the full VBA inventory is reconciled.

Artifact ledger: AZ trusted/frozen; BF functionally accepted/frozen; BG and BH failed/frozen; BI install/audit GREEN, manual acceptance PENDING.

## 2026-11-07 — BI operator UX acceptance findings; H8 dashboards prioritized

BI had green 33-test source gate and native installation/protection audit, but manual operator review identified major **product-level** gaps, so BI is **not UX accepted** and full worksheet hardening is paused. Requests: RefreshAll at most one terminal success dialog or remove if unnecessary; patient full names; explicit finalized-history coverage and period labels; prominent therapist workload and actual recorded absences with safe provenance; patient participation percentages and denominator; overall and per-person legible charts. Existing therapist daily/weekly/monthly/overall aggregators are implemented and may agree with only one finalized day; existing planned/actual utilization exists in wide table; absence source needs investigation. See detailed requirements, source findings, tests and safety gates in `Documentation/H8_STATISTICS_DASHBOARD_OPERATOR_FEEDBACK_2026-11-07.md`. Continue source investigation and H8 development before any new native smoke or broad protection. BI preserved unchanged as UX-gap evidence.
