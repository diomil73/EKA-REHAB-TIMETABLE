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

## Source checkpoint — 2026-11-07 H8 implementation (native gate pending)

Operator requested implementation without further approval. Tracked source commits have now changed the patient TSV and VBA renderer to show registry names and `Συμμετοχή %` (performed/scheduled), with strict duplicate-ID rejection, safe unknown-name fallback and zero denominator as empty; both patient outcome and therapist planned/actual slot overview native charts have been added to report VBA with single named chart deletion/rebuild on mode switch. Therapist report explains its finalized-day coverage. Existing related tests have been aligned to the changed schema/layout, and a new `Tests/test_h8_patient_names_and_participation.py` covers these source contracts.

The source checkpoint and exact next PowerShell source-test gate are in `Documentation/H8_STATISTICS_DASHBOARD_OPERATOR_FEEDBACK_2026-11-07.md` (latest section). **No tests have been reported after this source edit and no native H8 workbook exists.** BI is frozen diagnostic evidence, not an installation target. Further requirements still open: RefreshAll subordinate dialog diagnosis/remediation, full patient/therapist per-person charts, absence source/provenance, multi-period tests and acceptance. Do not claim H8 complete or authorize general protection.

## Operator's authoritative native RefreshAll dialog audit (2026-11-07)

The operator ran `Tests/test_h8_refresh_dialog_audit.py`: **3 passed in 0.68s**. Then the read-only tool inspected frozen BI and reported `ReadOnly: True`, `Workbook unchanged: True`, `H8 REFRESH DIALOG AUDIT COMPLETE`. The exact candidate success dialogs:
- Module3:17 `RefreshAll`: `MsgBox "Όλα ενημερώθηκαν.", vbInformation` (the only desired terminal success).
- Module1:380 `RefreshTherapistDaily`: `MsgBox "Το THERAPIST_DAILY ανανεώθηκε.", vbInformation`.
- Module2:229 `RefreshTherapistAttendanceAndConflicts`: `MsgBox "THERAPIST_ATTENDANCE και CONFLICT_LOG ανανεώθηκαν.", vbInformation`.
- Module4:217 `BuildReplacementSuggestions_V4`: `MsgBox "Οι προτάσεις αναπλήρωσης v4 δημιουργήθηκαν.", vbInformation`.
- Module4:336 nested `AddReplacementButtonsV4`: `MsgBox "Κουμπιά που δημιουργήθηκαν: " & createdCount, vbInformation`.
- Module10:68 `RestoreReturnedTherapists`: `MsgBox "Ολοκληρώθηκε ο έλεγχος επαναφοράς.", vbInformation`.

All six expected operational refresh procedures and nested call candidates were found; these are **candidate sites**, not proof each one triggers in every execution. The remaining non-message delegates `ColorInfectiousInMaster` and `ColorReplacementMarks` were found. Frozen BI was not edited.

## Source-only RefreshAll remediation checkpoint

New commits:
- `8b1f148` `Python/tools/install_h8_single_message_refresh.py`: pure preflight exact contract validation of original six-call `Module3.RefreshAll`; adds a public scoped flag `EKA_SilentRefreshSuccess`, one terminal success, one error handler with Excel session state restoration; wraps the five exact delegate-success dialogs only while the total refresh is underway, preserving independent success messages and other `vbCritical` messages. Fails closed if VBA procedure text or exact dialog occurrences differ. Installs only into a fresh disposable destination, followed by read-only persisted VBA source equality verification.
- `b4828fd` `Tests/test_h8_single_message_refresh.py`: source contracts and fail-on-drift, duplicate messages and reinjection.
- `46cbd0e`, `6debc43` `Python/tools/prepare_h8_dashboard_smoke.py`: one-shot fresh XLSM builder using trusted AZ via the existing H7 installer chain and then the H8 refresh installer. Does not patch frozen BI.
- `c496c82` `Tests/test_h8_dashboard_smoke_builder.py`: sequencing / refuse-overwrite / freeze-on-failure tests.

**Important:** These code changes have NOT passed a new operator test gate. None has been run in native Excel yet, and the H8 workbook has NOT been created. The existing 52-pass gate occurred before these commits. Do not claim verified single-dialog UX or broader dashboard acceptance before a new source test and native smoke. No change to historical store, clinic data, or frozen files was made.

### Next source gate

With Excel fully closed, from repo root:

```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }

python -m pytest `
  Tests/test_h8_single_message_refresh.py `
  Tests/test_h8_dashboard_smoke_builder.py `
  Tests/test_h8_refresh_dialog_audit.py `
  Tests/test_h8_patient_names_and_participation.py `
  Tests/test_patient_history_ui_h7.py `
  Tests/test_patient_history_reports_h7.py `
  Tests/test_h7_patient_mode_switch_and_toolbar_fix.py `
  Tests/test_therapist_history_ui.py `
  Tests/test_h7_integrated_toolbar_and_patient_order.py `
  Tests/test_h7_history_previous_screen_navigation.py -q

if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: H8 source regression tests failed" }
```

Only after this gate is green may a **fresh one-shot** candidate be created as `Excel/previews/UNIFIED_PATIENT_HISTORY_H8_20261107_BJ.xlsm` from frozen trusted `Excel/previews/UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm`, provided `BJ` does not already exist. The new builder is `python Python/tools/prepare_h8_dashboard_smoke.py --source ... --output ...`. After new build, repeat full independent `audit_workbook_protection.py`, `inspect_h7_integrated_ui.py`, and new `audit_h8_refresh_dialogs.py` (the latter lists lexical dialog candidates still present, intentionally under scoped flags). Freeze BJ on any failure; do not retry/patch an existing BJ. No manual interaction until audit passes. During native manual smoke, test RefreshAll **once** (six functions must execute), both reports across all four modes, names, KPI, aggregated and selected-person chart, mode switching, button placement, protection and return.

**Open H8 acceptance gaps**: final therapist absence provenance KPI and full longitudinal individual trend visualization remain to be built. First H8 chart release is still a trial, not full product acceptance. Full worksheet protection of the other 17 sheets remains unauthorized.

## Operator source gate GREEN — 2026-11-07; BJ native proof authorized

The operator confirmed the exact combined H8/H7 test command after the guarded RefreshAll installer and one-shot builder commits: **63 passed in 5.63s**. This supersedes the previous pending-test status for that source gate.

**Disposition:** H8 source regression gate GREEN. Authorize one fresh native candidate `Excel/previews/UNIFIED_PATIENT_HISTORY_H8_20261107_BJ.xlsm`, created only once from trusted frozen AZ using `Python/tools/prepare_h8_dashboard_smoke.py`. BI and earlier evidence remain untouched. This is **not** native VBA or chart success and not final H8 acceptance.

Exact one-shot PowerShell command (close Excel first):

```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }

$source = ".\\Excel\\previews\\UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm"
$output = ".\\Excel\\previews\\UNIFIED_PATIENT_HISTORY_H8_20261107_BJ.xlsm"
if (!(Test-Path -LiteralPath $source)) { throw "SAFETY STOP: AZ source missing" }
if (Test-Path -LiteralPath $output) { throw "SAFETY STOP: BJ already exists; do not overwrite or reuse" }

python .\\Python\\tools\\prepare_h8_dashboard_smoke.py --source $source --output $output
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BJ build failed; freeze BJ and do not retry" }

python .\\Python\\tools\\audit_workbook_protection.py $output
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BJ protection audit failed; freeze BJ" }

python .\\Python\\tools\\inspect_h7_integrated_ui.py $output
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BJ integrated toolbar audit failed; freeze BJ" }
```

Expected native build: H7 copy and installation sequence verified; `H8 REFRESH ONE-MESSAGE INSTALL OK`, `Persisted VBA reopen verification: OK`, `H8 DASHBOARD PREPARATION OK`. Protection audit must report both patient and therapist report `ProtectContents=True`, `ProtectDrawingObjects=False`; all four technical history sheets protected/veryHidden; `ReadOnly=True` and `Workbook unchanged: True`. Toolbar audit should confirm canonical RefreshAll and two statistics buttons, no overlaps or missing macros. No manual button clicks or Save/Exit until the complete read-only output is reviewed.

If any step fails, do not re-run against the same BJ; report full terminal output. If BJ existed before this command, stop at the safety gate and preserve it. After audit GREEN, manually smoke test native charts, report modes and one terminal RefreshAll success only (failure messages retained). Charts and success-dialog behavior have **not** yet passed native proof.

Open H8 scope remains genuine therapist absence provenance, clinical denominator policy, and longitudinal individual trends, as well as broad worksheet protection. Do not conflate source green with delivery acceptance.

## 2026-11-07 — H8 BJ native creation and read-only audits GREEN; manual UX pending

Operator evidence confirms fresh BJ (`UNIFIED_PATIENT_HISTORY_H8_20261107_BJ.xlsm`) from frozen AZ was built successfully, including H7 integrated installation and H8 one-message RefreshAll install; persisted VBA reopen verification OK. Independent workbook protection audit: 23 sheets; both reports contents-protected/drawing objects unprotected; four history technical sheets protected+veryHidden; 17 baseline sheets still unprotected; `ReadOnly=True`, `Workbook unchanged=True`, audit complete. Independent integrated toolbar audit: canonical RefreshAll and both statistics buttons aligned and macros present, `H7 INTEGRATED AUDIT OK; workbook not changed`. Previous source gate was `63 passed in 5.63s`. **No actual manual chart or refresh-click result yet; do not mark H8 accepted or run broader sheet hardening.** BJ now authorized for one controlled manual smoke (reports, overview/selected-person charts, names, completion %, period switches/navigation, one RefreshAll click); freeze on any anomaly. See latest section of `Documentation/CURRENT_HANDOFF.md` for exact state, stops, and artifact ledger.

## 2026-11-07 — BJ manual acceptance GREEN; H8 visual refinement source staged

The operator explicitly reported **«όλα λειτουργούν»** after native manual checks of BJ `UNIFIED_PATIENT_HISTORY_H8_20261107_BJ.xlsm`. This is the affirmative operator result for the previously requested BJ report/chart/mode/navigation/RefreshAll smoke. Together with prior BJ native install/read-only protection/integrated UI audits, freeze **BJ as native manually accepted baseline**. Do not edit, overwrite or reinstall BJ. This acceptance applies to the BJ functionality that was actually tested; it does not mean future refinements or all outstanding analytics scope are finished.

The operator then specified a new follow-up slice:

- Daily/weekly/monthly/overall report buttons must show a clearly distinct **selected active mode** color.
- Therapist patient-count scale: **6 red, 5 orange, 3–4 yellow, 1–2 green**, and 0 neutral gray; other than daily, use average **distinct patients per finalized therapist-day** with clear non-integer thresholds (>=5.5 red, >=4.5 orange, >=2.5 yellow, >0 green). Do not confuse delivered sessions with distinct patients or occupied time slots. Report remains version-safe CURRENT_FINAL only.
- **Three compact same-row therapist panels**: actual patients/day on a 0–6 reference scale; scheduled-vs-completed treatment sessions; and therapist absences vs working days. All should be clearly visible near the top.
- Patient overview/selected-person chart should stay visible when scrolling to rows near the end of the long patient list.

Tracked source implementation since BJ acceptance:
- `c7ec124`: patient active-mode button colors (dark-blue selected, light inactive) and freeze top 22 rows so chart remains visible while patient table scrolls.
- `873e55e`, `26c4d6b`: therapist active-mode colors, 3 native compact dashboard panels in columns A:D / E:H / I:M, load traffic light, 0–6 reference axis, one named chart each, refresh without chart stacking, chart captions legible.
- `e18d020`: therapist TSV exporter now derives distinct delivered patient IDs per therapist/date from `read_current_patient_session_history_rows` (CURRENT_FINAL), protects against duplicate current-final session keys, calculates average across therapist finalized reported days; adds `UniquePatientsPerDay` column. This deliberately distinguishes unique patients from sessions and occupied slots.
- `55e6a84`, `685deed`: workbook renderer displays and traffic-colors `Ασθενείς/ημέρα` in column N and uses this derived KPI for the 0–6 first plot.
- `f4c8ade`, `bed39a2`, `e71c431`: aligned affected existing regression source tests with new 15-column TSV and sticky chart.
- `06061b1`, `17346ed`: new focused source tests for mode selection, sticky chart, 3-panel layout, distinct patient counts, failure safety.

**Important limitation:** The current finalized therapist/day workload schema does not persist authentic dated therapist absence/working-day facts. The append-only `DAY_HISTORY_EVENTS` currently enumerates only finalization/amendment event types, not daily staff attendance. The **third chart is deliberately an empty-data panel titled «Παρουσίες / απουσίες» with an explicit visible label «Μη διαθέσιμο ιστορικό απουσιών / εργασίμων»**. Never show missing data as 0 absences, never infer days absent from gaps in snapshots or from current editable `THERAPIST_ATTENDANCE`. Authentic absence days / expected-working-day history require a separate durable finalized source and agreed holiday/leave policy. Do not describe third panel as a populated quantitative chart until that source exists.

**Status:** BJ manually GREEN and frozen; new visual-refinement source changes committed but **NOT TESTED OR NATIVE INSTALLED**. New one-shot workbook must use a fresh suffix (e.g. BK) from trusted AZ using existing H8 builder; must not overwrite BJ. Broad worksheet protection remains unauthorized.

### Next source gate (Excel closed)

```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }

python -m pytest `
  Tests/test_h8_dashboard_visual_navigation.py `
  Tests/test_h8_unique_patient_workload.py `
  Tests/test_h8_single_message_refresh.py `
  Tests/test_h8_dashboard_smoke_builder.py `
  Tests/test_h8_refresh_dialog_audit.py `
  Tests/test_h8_patient_names_and_participation.py `
  Tests/test_patient_history_ui_h7.py `
  Tests/test_patient_history_reports_h7.py `
  Tests/test_h7_patient_mode_switch_and_toolbar_fix.py `
  Tests/test_therapist_history_ui.py `
  Tests/test_h7_integrated_toolbar_and_patient_order.py `
  Tests/test_h7_history_previous_screen_navigation.py -q

if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: H8 visual-refinement source gate failed" }
```

If source gate fails, fix source/tests before building anything. On green, create one fresh BK from AZ via `Python/tools/prepare_h8_dashboard_smoke.py`, full read-only protection and H7 integrated UI audits before opening, then one native manual smoke of active-mode colors, chart placements/sticky scrolling, unique patient workload/color thresholds, selected person chart, report protection and RefreshAll. Freeze any failed BK without repeat/repair. Source tests do not prove VBA chart COM or Excel freeze-pane behavior.

## H8 visual source-test regression failure and fix — 2026-11-07 late evening

Operator ran the prescribed 12-file visual regression suite and reported **71 passed, 1 failed in 7.17s**. The failing assertion in `Tests/test_therapist_history_ui.py::test_therapist_report_has_idempotent_workload_chart_above_history_table` expected the old single-chart text `series.Values = Array(plannedSlots, actualSlots)`. That expression was intentionally replaced in H8: the first of three native charts now uses `series.Values = Array(loadPerDay)` based on unique patients/day and a traffic-light color, while the second uses `series.Values = Array(plannedSessions, actualSessions)`; the third is absence history explicitly unavailable. **No evidence of a runtime failure, source-code failure or native Excel failure from this assertion.** PowerShell subsequently threw the expected `SAFETY STOP: H8 visual tests failed`. This stop is required and prevents creating BK.

Source fix commit **`61c3c99`** updates only the stale test assertion to check both new chart payloads, color mapping and third panel removal contract; it does not change VBA or the frozen BJ.

**Current status:** FIX COMMITTED, **green rerun pending**. Do not claim 72 pass or build BK until a fresh operator gate succeeds.

Next, Excel closed, run:

```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: git pull failed" }

python -m pytest Tests/test_therapist_history_ui.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: therapist regression still failed" }

python -m pytest `
  Tests/test_h8_dashboard_visual_navigation.py `
  Tests/test_h8_unique_patient_workload.py `
  Tests/test_h8_single_message_refresh.py `
  Tests/test_h8_dashboard_smoke_builder.py `
  Tests/test_h8_refresh_dialog_audit.py `
  Tests/test_h8_patient_names_and_participation.py `
  Tests/test_patient_history_ui_h7.py `
  Tests/test_patient_history_reports_h7.py `
  Tests/test_h7_patient_mode_switch_and_toolbar_fix.py `
  Tests/test_therapist_history_ui.py `
  Tests/test_h7_integrated_toolbar_and_patient_order.py `
  Tests/test_h7_history_previous_screen_navigation.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: H8 visual source regression still failed" }
```

If any failure persists, freeze the next native step and investigate the new traceback, without retrying or overwriting existing frozen XLSM. Only a green source gate authorizes fresh uniquely named BK native smoke and independent protection/integrated audits. Previous BJ manual pass remains valid and frozen.

## Operator H8 visual-regression source gate GREEN — 2026-11-07

The operator reran tests after commit `61c3c99` fixed the obsolete single-chart assertion:
- Focused `Tests/test_therapist_history_ui.py`: **17 passed in 2.76s**.
- Complete H8 visual/H7 12-file regression suite: **72 passed in 5.03s**.

This resolves the preceding **71 passed, 1 failed** source-only gate. The active-mode colors, CURRENT_FINAL distinct-patient/day color bands, therapist 3-panel renderer, sticky patient chart and earlier H8 source regressions now pass the specified tests. **This is source-test evidence, not native Excel behavior.** The old BJ remains frozen and manually accepted. No BK build result is reported yet.

**Authorized next native step:** from trusted frozen `Excel/previews/UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm` create exactly one fresh `Excel/previews/UNIFIED_PATIENT_HISTORY_H8_20261107_BK.xlsm` via `Python/tools/prepare_h8_dashboard_smoke.py` with Excel closed, refusal if BK already exists, no retries on any failed native install. After success: `audit_workbook_protection.py` and `inspect_h7_integrated_ui.py` read-only gates. If all green, operator tests four active-mode colors, 1–2 green / 3–4 yellow / 5 orange / 6+ red using distinct patient counts and period coverage, three side-by-side panels (absence panel must state unavailable, not invent zero), sticky patient overview when selecting final rows, selected-person charts, clickability/protection and one-message RefreshAll. If any check fails, freeze BK as diagnostic and report exact step/message. Do not overwrite AZ/BJ/BK or modify production. Authentic therapist working-days/absence metrics require a new durable source and are not silently inferred.

Exact next build with Excel closed:

```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }

$source = ".\\Excel\\previews\\UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm"
$output = ".\\Excel\\previews\\UNIFIED_PATIENT_HISTORY_H8_20261107_BK.xlsm"
if (!(Test-Path -LiteralPath $source)) { throw "SAFETY STOP: AZ source missing" }
if (Test-Path -LiteralPath $output) { throw "SAFETY STOP: BK already exists; do not overwrite" }

python .\\Python\\tools\\prepare_h8_dashboard_smoke.py --source $source --output $output
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BK build failed. Freeze BK and do not retry" }

python .\\Python\\tools\\audit_workbook_protection.py $output
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BK protection audit failed. Freeze BK" }

python .\\Python\\tools\\inspect_h7_integrated_ui.py $output
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BK toolbar audit failed. Freeze BK" }
```

## Operator BK native build and audits GREEN — 2026-11-07

The operator pasted the full native PowerShell transcript for `Excel/previews/UNIFIED_PATIENT_HISTORY_H8_20261107_BK.xlsm`, fresh from trusted frozen AZ and after prior **17 passed / 72 passed** regression gates. The transcript was provided as an attachment when the UI did not display the preceding response.

Native one-shot installation all successful: `H7 INTEGRATED COPY VERIFIED`, `UNIFIED DAY SAVE UI INSTALL OK`, `THERAPIST HISTORY UI INSTALL OK`, `H7 PATIENT HISTORY UI INSTALL OK`, `H7 INTEGRATED UI INSTALL OK`, `H8 REFRESH ONE-MESSAGE INSTALL OK`, `H8 DASHBOARD PREPARATION OK`, and `Persisted VBA reopen verification: OK`. No source or frozen evidence was changed.

Independent read-only protection audit reports **23 worksheets**, both `PATIENT_HISTORY_REPORT` and `THERAPIST_HISTORY_REPORT` `ProtectContents=True`, `ProtectDrawingObjects=False`, `ProtectScenarios=True`; four finalized-history technical sheets `Visible=2` and fully protected; 17 baseline sheets remain unprotected (expected, general protection gate paused), `ReadOnly=True`, `Workbook unchanged: True`, `WORKBOOK PROTECTION AUDIT COMPLETE`.

Independent integration audit: `ReadOnly=True`, `MASTER SaveOperationalDay` present, canonical `RefreshAll` button left=170, top=563.6, caption `ΑΝΑΝΕΩΣΗ`; both statistics buttons at left=270 and 424 on same row, correct macro bindings, all four expected VBA modules present, both reports visible, technical history veryHidden; `H7 INTEGRATED AUDIT OK; workbook not changed`. No audit failure or `SAFETY STOP`.

**Disposition:** BK native creation/installation/persisted verification/protection/toolbar audits GREEN, **manual Excel functional UI acceptance pending**. BJ stays separately frozen and accepted. Do not re-run installer or overwrite BK.

Next authorized native manual test: open BK only with macros enabled, check selected-mode color across all four modes in both reports; patient overview remains visible while scrolling to bottom, individual selected-row graph; therapist three compact side-by-side panels, distinct-patient/day color thresholds (1–2 green, 3–4 yellow, 5 orange, 6+ red) on report samples, planned vs actually performed therapies; third absence panel must explicitly report missing historical attendance/working-day data, not falsely display zero; back navigation; protected report buttons still clickable; RefreshAll once and terminal message only. Do not finalize/amend day or change production. Stop and freeze BK diagnostic on any error/unexpected metric, save output/screenshot, no reinstall/retry. Missing therapist-absence historical source is a known incomplete functional requirement, not a failure of BK installation.

## 2026-11-07 — BK operator manual feedback and follow-up source fix

Operator reported BK **«κατά τα υπόλοιπα όλα δείχνουν σωστά»**, with two precise visual changes requested:
1. Patient charts (both overall and selected patient) must show exactly **Προγραμματισμένες, Πραγματοποιημένες, Απουσίες**. Prior version instead showed completed, absences, cancellations. Cancellations remain in the patient history table and denominator; they are not erased or confused with patient absences.
2. The therapist load chart title should not contain **«(0-6)»**. Keep the numeric axis scale and traffic-light thresholds unchanged; use **«Ασθενείς ανά ημέρα»**.

Changes committed after BK:
- `273396a`: patient VBA aggregates scheduled=TSV fields(5), performed=fields(6), absences=fields(7), and renders all three categories for the overview. Selected-row chart maps columns E/F/G to the same categories. Does not alter underlying history or cancellation data.
- `4ae2882`: therapist chart title `Ασθενείς ανά ημέρα`; actual scale remains.
- `fb85c82`, `bc6177c`: regression source contracts updated and two focused expectations added.

**Status**: BK native install/protection/navigation audits GREEN; operator manual feedback is *generally positive*, but BK does not satisfy the two revised visualization labels/data. Preserve BK as frozen manual feedback artifact and BJ as previously accepted frozen baseline. **No new source-test results since these changes**, no next native workbook built. Future real data may reveal the need for further adjustment. Third therapist attendance panel remains marked unavailable, until trustworthy working-day/absence history exists. Do not label it as populated.

Next, with Excel closed:
```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }

python -m pytest `
  Tests/test_h8_dashboard_visual_navigation.py `
  Tests/test_h8_unique_patient_workload.py `
  Tests/test_h8_single_message_refresh.py `
  Tests/test_h8_dashboard_smoke_builder.py `
  Tests/test_h8_refresh_dialog_audit.py `
  Tests/test_h8_patient_names_and_participation.py `
  Tests/test_patient_history_ui_h7.py `
  Tests/test_patient_history_reports_h7.py `
  Tests/test_h7_patient_mode_switch_and_toolbar_fix.py `
  Tests/test_therapist_history_ui.py `
  Tests/test_h7_integrated_toolbar_and_patient_order.py `
  Tests/test_h7_history_previous_screen_navigation.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: H8 patient visualization regression failed" }
```

Only if this gate passes, build one fresh new native preview with a previously unused suffix (e.g. `BL`) from immutable trusted AZ using the one-shot H8 builder. Run read-only protection and integrated UI audit on BL before manual functional use. On any failure freeze BL, do not reinstall or overwrite. Preserve all historic proofs.

## 2026-11-07 — operator source gate 74 GREEN; one-shot BL native smoke next

The operator reported **74 passed in 4.55s** from the 12-file H8/H7 regression gate run after commits `273396a` (patient charts now scheduled/performed/absences), `4ae2882` (therapist load title omits "(0-6)"), `fb85c82` and `bc6177c` (updated regression tests). Thus this specific source gate is GREEN, not native chart/UX proof.

Preserve frozen BJ (previously manually accepted) and BK (native audits and mostly positive UX feedback, two subsequent visual changes requested). Neither should be rebuilt, repaired, or overwritten.

Authorize exactly one new native preview **`Excel/previews/UNIFIED_PATIENT_HISTORY_H8_20261107_BL.xlsm`**, from trusted frozen `Excel/previews/UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm` using `Python/tools/prepare_h8_dashboard_smoke.py`, with Excel fully closed. Stop if source missing or BL already exists. After one successful creation, independently audit with `audit_workbook_protection.py` and `inspect_h7_integrated_ui.py`, both read-only. On *any* failure, freeze BL and do not retry/reinstall it; no manual Excel interaction before audits green.

Expected native state mirrors BK: both report sheets contents-protected with drawing objects unlocked; all four finalized-history technical sheets protected and veryHidden; only existing 17 sheets remain outside broad protection; read-only reports and integrated toolbar present. Once native green, manually test patient chart triple values in both overall and selected patient, new therapist title with unchanged scale and color bands, active period buttons, 3 compact therapist panels with absence history explicitly unavailable, sticky patient chart while scrolling, navigation and one terminal RefreshAll success. Do not falsify working-day or absence metrics. Preserve clinical history and source of truth; mark BL manually accepted only on actual operator observation.

## Operator BL native creation and read-only audits GREEN — 2026-11-07

The operator provided full terminal output for `Excel/previews/UNIFIED_PATIENT_HISTORY_H8_20261107_BL.xlsm`, created from frozen trusted AZ after the prior **74 passed in 4.55s** H8 source regression gate. This is an authoritative native installer/audit result, not yet manual UX acceptance.

- Build: `H7 INTEGRATED COPY VERIFIED`, `UNIFIED DAY SAVE UI INSTALL OK`, `THERAPIST HISTORY UI INSTALL OK`, `H7 PATIENT HISTORY UI INSTALL OK`, `H7 INTEGRATED UI INSTALL OK`, `H8 REFRESH ONE-MESSAGE INSTALL OK`, `Persisted VBA reopen verification: OK`, and `H8 DASHBOARD PREPARATION OK`. Both old source/evidence and frozen BJ/BK remain untouched.
- Independent read-only worksheet protection audit: `ReadOnly: True`, `Worksheet count: 23`, `Workbook unchanged: True`, `WORKBOOK PROTECTION AUDIT COMPLETE`. Both `PATIENT_HISTORY_REPORT` and `THERAPIST_HISTORY_REPORT` are visible and `ProtectContents=True`, `ProtectDrawingObjects=False`, `ProtectScenarios=True`; four history sheets `FINALIZED_DAY_SNAPSHOTS`, `THERAPIST_DAY_HISTORY`, `PATIENT_SESSION_HISTORY`, and `DAY_HISTORY_EVENTS` are `Visible=2` (veryHidden), fully protected. The remaining 17 sheets continue at the prior unprotected baseline. Their audit `SAFETY FINDING` lines reflect the previously deferred general protection project, **not a newly failed H8 gate**.
- Independent integrated toolbar audit: `ReadOnly: True`, canonical `RefreshAll` caption `ΑΝΑΝΕΩΣΗ` on `THERAPIST_DAILY`, therapist/patient report buttons aligned, `Module3`, `modUnifiedDaySave`, `modTherapistHistoryReport`, `modPatientHistoryReport` present; reports visible and history sheets veryHidden; `H7 INTEGRATED AUDIT OK; workbook not changed`.

**Disposition:** BL native creation, persisted VBA verification, protection and toolbar audits GREEN. **Manual Excel testing is not yet reported**. BL must remain single-use, no installer retry, overwrite or repair. Open BL with macros enabled solely for controlled manual testing: patient overall and selected-patient chart shows `Προγραμματισμένες / Πραγματοποιημένες / Απουσίες`, therapist load chart title `Ασθενείς ανά ημέρα` without the text `(0-6)`, unchanged traffic scale and three compact panels (therapist absence/working-day panel intentionally says missing history), four active-mode colors, sticky patient chart with late-list rows, both report protections, navigation and one final `RefreshAll` message. Stop and freeze diagnostic on any failure, record exact symptom; no data/day close or production changes. BJ and BK remain frozen.

Current H8 unfinished clinical-data scope: real therapist-absence-to-working-day ratio requires durable provenance/leave/holiday facts. Source tests and no-history placeholder do not establish real absence measurements.

## 2026-11-08 — BL manual operator acceptance GREEN

The operator explicitly confirmed **«όλα καλά»** after being asked to manually verify `Excel/previews/UNIFIED_PATIENT_HISTORY_H8_20261107_BL.xlsm` in native Excel. This confirms the previously pending BL manual smoke, in addition to its earlier successful build/persisted VBA and independent read-only protection and integrated-toolbar audits, and the 74 passed in 4.55s source regression gate.

**Disposition: BL ACCEPTED for the tested H8 dashboard UI scope and FROZEN.** The accepted scope includes the new patient chart categories **Προγραμματισμένες / Πραγματοποιημένες / Απουσίες** in overview and selected-patient views, cleaned therapist chart title **Ασθενείς ανά ημέρα**, active period-button styling, sticky patient chart scrolling, side-by-side therapist panels, navigation, report protection and single terminal refresh success. This is the operator's concise acceptance, not an independently captured click-by-click replay; do not overstate deeper longitudinal data validation.

**Do not modify/rebuild/reinstall BL, BK or BJ.** Trusted AZ remains immutable. No new workbook or broad worksheet-protection roll-out is authorized by this acceptance alone. Any future change requires new suffix, source tests, fresh one-shot build, independent audits and operator acceptance.

**Open follow-up (not included in BL completion):** historical therapist presence/absence against valid working days remains unavailable from the current finalized therapist/day schema. Third chart intentionally identifies missing data instead of inventing zero absences. Future work must first design and persist authentic therapist-day availability/absence facts, working-day/holiday policies, reason provenance, correction/amendment behavior, and tests before drawing quantitative attendance ratios. Current reports should be revisited as more real finalized-day data accumulate; no new requirements have yet been confirmed.

## 2026-11-08 — H9 historical therapist attendance domain contract started

After operator accepted and froze H8 BL and agreed to continue, a read-only review confirmed that existing `THERAPIST_DAY_HISTORY`, `FINALIZED_DAY_SNAPSHOTS` and `DAY_HISTORY_EVENTS` do not store proven therapist presence/absence facts with valid expected-working-day denominators. The live `THERAPIST_ATTENDANCE` projection must not be treated as immutable history, and missing sessions are not equivalent to a staff absence.

New source-only commits:
- `0a77e13`: `Python/rehab_core/therapist_attendance_history.py` pure validated attendance facts and fail-closed summary requiring explicit verified calendar, presence/absence evidence, reason, source provenance and current-final snapshot link.
- `a0ac130`: seven focused pure-domain unit tests covering valid numerator/denominator, missing dates, superseded versions, duplicate days, missing provenance/reason, nonworking-day exclusion, no-data ratio.
- `f98ddad`: `Documentation/H9_THERAPIST_ATTENDANCE_HISTORY_CONTRACT.md` explains audited limitations, completion dependencies and the next source gate.

**Status:** contract committed but **NOT TESTED by operator**; no runtime Excel changes, no H9 installer, no new native preview. BL/BK/BJ and trusted AZ remain frozen/unchanged. Attendance chart still honestly shows missing data. Operator next:
```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
python -m pytest Tests/test_therapist_attendance_history_contract.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: H9 attendance-domain source gate failed" }
```
Next implementation step only after source green and explicit business rules for scheduled working days, holiday and leave handling, provenance, versioned persistence and amendment. No H9 native workbook is authorized by the domain test alone.

## 2026-11-08 — H9 attendance contract source gate GREEN

Operator ran the first focused H9 pure-domain source gate after commit `a0ac130` and reported **7 passed in 0.98s**. This validates only the model/contract: explicit attendance and expected-working-date facts, provenance, absence reasons, duplicate and missing dates, superseded versions and nonworking-day safeguards.

**Current authorized next development:** Implement a strictly read-only workbook attendance history schema reader that rejects missing/malformed source data (never treats absent sheet as zero absences) and validates links against CURRENT_FINAL snapshot IDs. Do not modify Excel files or create a new H9 preview. Storage writes, actual roster-calendar provenance, holiday/leave policy and close-day integration remain separately gated. The accepted BL remains frozen. Next check must cover both the seven green contract tests and any new read-only-store tests, with no claim of native Excel proof.

## 2026-11-08 — H9 source-only read-only historical attendance layer, operator tests pending

Operator confirmed initial H9 domain gate **7 passed in 0.98s**. Continuing autonomously, added:
- `cb3c652`: read-only `Python/rehab_excel/therapist_attendance_history_store.py` for future version-linked `THERAPIST_ATTENDANCE_HISTORY` schema, with strict missing-source safety, CURRENT_FINAL selection, mandatory independent expected-working-date calendar, and no workbook writes.
- `11b392b`: six synthetic-workbook regression tests for source absence, immutable reads, superseded corrections, missing evidence, invalid source and duplicates.
- `c5b7d85`: expanded `Documentation/H9_THERAPIST_ATTENDANCE_HISTORY_CONTRACT.md`.

**Status:** New source files committed, **new six tests have not been executed by the operator**. Run:
```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
python -m pytest Tests/test_therapist_attendance_history_contract.py Tests/test_therapist_attendance_history_store.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: H9 attendance history read-only gate failed" }
```
No H9 native Excel workbook, sheet writer, chart data, or generalized protection changes yet. The accepted BL stays frozen. If the gate fails, fix source and rerun; never patch a frozen workbook. If green, next task is explicit finalized working-day roster and attendance provenance/write-path design, with calendar/leave business rules, before any UI figures can be trusted.

## 2026-11-08 — H9 read-only gate GREEN, fresh H8 illustrative history generator staged

Operator confirmed **13 passed in 8.33s** for \`Tests/test_therapist_attendance_history_contract.py\` + \`Tests/test_therapist_attendance_history_store.py\`. This supersedes their previous pending gate. H9 domain/read-only source gate is GREEN, **not** real Excel attendance capture.

New request: populate the entirely synthetic/random-name Excel test environment with varied **multiweek historical data** for compelling end-to-end daily/weekly/monthly/overall chart tests, without modifying any accepted frozen Excel.

Commits:
- \`b415d25\`, \`7174b19\`: \`Python/tools/seed_h8_synthetic_history.py\`, deterministic synthetic history for 2026-09-21..2026-10-23 weekdays; existing patient/therapist IDs; internally consistent CURRENT_FINAL snapshots, per-session outcomes, therapist workload, append-only finalization events and H9 synthetic roster facts. Excel COM bulk write to fresh \`_DEMO_\` workbook, read-only reopen verification, stop-on-mismatch.
- \`41097f9\`: \`Python/tools/prepare_h8_synthetic_demo.py\` one-shot wrapper using existing H8 native installer then seed, output refusal/no retry.
- \`53ec37e\`: \`Tests/test_h8_synthetic_demo_history.py\`, source-only deterministic/multiperoid reconciliation, absence and identifier integrity and filename safety tests.
- \`8752118\`: \`Documentation/H8_SYNTHETIC_HISTORY_DEMO_PLAN.md\`, complete command procedures/contingencies and explicit demo-only data semantics.

**Current state:** Generator/source tests committed but **NOT YET EXECUTED**, and **no BM workbook created**. BL/BK/BJ/AZ untouched and frozen. Seeded H9 staff attendance is expressly illustrative; third H8 therapist graph remains an unavailable-source placeholder and is not quietly populated from a demo sheet.

The prescribed next source-only gate (Excel closed):
\`\`\`powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
python -m pytest Tests/test_h8_synthetic_demo_history.py Tests/test_therapist_attendance_history_contract.py Tests/test_therapist_attendance_history_store.py Tests/test_h8_dashboard_visual_navigation.py Tests/test_h8_unique_patient_workload.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Synthetic demo source tests failed" }
\`\`\`
After source GREEN only, create a **single** fresh \`Excel/previews/UNIFIED_H8_DEMO_HISTORY_20261108_BM.xlsm\` from trusted AZ using \`prepare_h8_synthetic_demo.py\`; do not rerun/reinstall/overwrite on any failure, freeze that output and pick a different unused suffix after code-level correction. Follow with read-only \`audit_workbook_protection.py\` and \`inspect_h7_integrated_ui.py\`, then manually test multi-period visualization. Complete operator recipe and expected 24-sheet audit in the specific plan. Do not infer native success from source tests.

## 2026-11-08 — synthetic multiweek history source gate GREEN

The operator ran the prescribed five-file synthetic-history/H9/H8 source regression gate after commits `b415d25`, `7174b19`, `41097f9`, `53ec37e`, and confirmed **30 passed in 3.70s**. Thus the deterministic cross-month fixture, in-memory history integrity checks, guarded builder and H9 pure/read-only contracts pass the requested source test selection. **This is not native Excel proof** and there has not yet been a native BM workbook build or chart/manual review.

**Authorize ONE fresh candidate:** `Excel/previews/UNIFIED_H8_DEMO_HISTORY_20261108_BM.xlsm` based on trusted frozen `Excel/previews/UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm`, with all Excel windows closed. The one-shot `Python/tools/prepare_h8_synthetic_demo.py` builds existing H8 native UI then seeds synthetic CURRENT_FINAL days plus H9 demo attendance data. Never overwrite or retry on a failed BM, and never change AZ/BJ/BK/BL or production.

Native sequence:
```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }

$source = ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm"
$demo = ".\Excel\previews\UNIFIED_H8_DEMO_HISTORY_20261108_BM.xlsm"
if (!(Test-Path -LiteralPath $source)) { throw "SAFETY STOP: AZ source missing" }
if (Test-Path -LiteralPath $demo) { throw "SAFETY STOP: BM already exists; no overwrite" }

python .\Python\tools\prepare_h8_synthetic_demo.py --source $source --output $demo
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BM native seed failed; freeze BM and do not retry" }

python .\Python\tools\audit_workbook_protection.py $demo
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BM protection audit failed; freeze BM" }

python .\Python\tools\inspect_h7_integrated_ui.py $demo
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BM toolbar audit failed; freeze BM" }
```

Expected additional `THERAPIST_ATTENDANCE_HISTORY` demo technical sheet means 24 sheets, not BL's 23. The H8 third therapist absence/working-day chart is **not integrated**; its honest unavailable-state remains. For BM native build SUCCESS expect `H8 SYNTHETIC HISTORY SEED OK`, `Persisted current-final readback: OK`, `H8 SYNTHETIC DEMO PREPARATION OK`; if any check fails, freeze with full logs and do not repeat on same BM. Only once both independent read-only audits pass may operator open the demo and verify distinct periods and visual charts. The seed is DEMO ONLY and must never be transferred to production.
