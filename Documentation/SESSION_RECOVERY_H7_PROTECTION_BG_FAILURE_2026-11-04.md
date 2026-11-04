# Session Recovery — H7 therapist-report protection / failed BG smoke
Last updated: 2026-11-04
Repository: diomil73/EKA-REHAB-TIMETABLE
Active branch: work/history-statistics-rebuild

> Compact recovery packet for a new ChatGPT conversation. Read this first, then the linked specialist notes. It supersedes older next-step instructions where they conflict.

## 1. Exact project state

### Accepted / frozen
- H7 patient and therapist history/statistics UI is functionally accepted by the operator on disposable workbook **BF**.
- BF: Excel/previews/UNIFIED_PATIENT_HISTORY_H7_20261104_BF.xlsm.
- BF integrated read-only audit passed; operator manually confirmed toolbar placement, retained RefreshAll, patient and therapist reports, and direct previous-screen navigation.
- BF is frozen green evidence. Never patch, overwrite, or reuse it as an install destination.
- Trusted source for the next fresh integrated smoke: **AZ**, Excel/previews/UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm. Do not modify it.

### Failed / frozen
- **BG failed** during integrated preparation: Excel/previews/UNIFIED_PATIENT_HISTORY_H7_20261104_BG.xlsm.
- Preparation reached therapist-history UI installation and stopped with: SAFETY STOP: Therapist-history report buttons must remain clickable.
- BG is failed evidence. Do not open to edit, repair, rerun against, overwrite, or reuse it. Do not delete it as a substitute for preserving the evidence.
- AZ and BF remain unchanged.

### Current active work
Only the narrow protection of the blank visible THERAPIST_HISTORY_REPORT sheet is being addressed. Do not apply blanket protection to the other unprotected sheets yet.

BF read-only inventory:
- 23 worksheets total.
- 18 with ProtectContents=False; five protected.
- The four technical history sheets are protected and veryHidden.
- PATIENT_HISTORY_REPORT is protected.
- THERAPIST_HISTORY_REPORT was visible, blank and unprotected.
- UsedRange.Locked=True on all sheets does not enforce protection while ProtectContents=False.
- Workbook structure/windows are not protected. Do not enable structure protection blindly because workflows create/hide sheets.

## 2. Latest code corrections and commits

Original targeted protection fix:
- 45ed8c9: tests for report protection after install and error cleanup.
- 9a1574e: report VBA cleanup re-protects on error; installer protects the blank report while keeping drawing objects unprotected.

Persisted verification:
- 5c3aaea: save, close, reopen in separate Excel session with ReadOnly=True, then verify persisted report state.
- 870c1db: regression test for persisted read-only verification.

BG failure follow-up:
- 03f6afb: explicit positional COM protection call report_ws.Protect("", False, True, True); explicit empty-password Unprotect; verify pre/post state; report actual protection flags in error.
- d8e8418: persisted audit checks all five report buttons are visible and have an OnAction binding; includes actual ProtectContents / ProtectDrawingObjects values in the failure message.
- 27339cf, 5beda80: update test assertions to match the new helper/COM call.
- fd551a3, 5c766ff: ensure all four technical history sheets are checked for veryHidden state, including PATIENT_SESSION_HISTORY.

The branch also contains documentation updates recording the failure and next gate. At the time this recovery packet was prepared, branch HEAD was 54f4659791ce8c81ae25a456da3a2f0bec045f73; this documentation commit will advance HEAD. Always pull the branch rather than relying on a remembered SHA.

## 3. Test status: NOT YET GREEN AFTER THE LATEST FIX

The latest focused test file contains **15 tests**. The user previously ran the 13-test version after an earlier correction and got 12 passed / 1 failed; that failure was a stale assertion looking for technical-sheet names in the installer entry point rather than its _audit_persisted helper. That test was corrected.

The later BG preparation was attempted despite the preceding failed test gate, and failed at the report-button clickability assertion. The code and tests were then changed again. **No result has yet been supplied for the current 15-test version.**

Immediate next action, with Excel closed:

    git pull --ff-only origin work/history-statistics-rebuild
    if ($LASTEXITCODE -ne 0) {
        throw "SAFETY STOP: Git pull failed"
    }

    python -m pytest Tests/test_therapist_history_ui.py -q
    if ($LASTEXITCODE -ne 0) {
        throw "SAFETY STOP: Therapist history tests failed"
    }

Expected: 15 passed. If not, stop and diagnose. Do not run any smoke-preparation command after a failed test gate.

## 4. Only after all 15 tests pass: create a new BH

Use the trusted AZ source and a new destination. Never use BG as a source. Never use BF as an install destination.

    $source = ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm"
    $output = ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H7_20261104_BH.xlsm"

    if (!(Test-Path -LiteralPath $source)) {
        throw "SAFETY STOP: Trusted AZ source missing"
    }
    if (Test-Path -LiteralPath $output) {
        throw "SAFETY STOP: BH already exists. Do not overwrite."
    }

    python .\Python\tools\prepare_h7_integrated_ui_smoke.py --source $source --output $output
    if ($LASTEXITCODE -ne 0) {
        throw "SAFETY STOP: BH preparation failed. Freeze BH; do not retry or repair."
    }

If BH preparation succeeds, run the read-only audit:

    python .\Python\tools\audit_workbook_protection.py ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H7_20261104_BH.xlsm"
    if ($LASTEXITCODE -ne 0) {
        throw "SAFETY STOP: BH protection audit incomplete"
    }

Expected targeted delta:
- THERAPIST_HISTORY_REPORT.ProtectContents=True.
- THERAPIST_HISTORY_REPORT.ProtectDrawingObjects=False, so the five report shapes remain clickable.
- All five report buttons visible and bound to non-empty OnAction macros.
- THERAPIST_HISTORY_REPORT remains visible.
- FINALIZED_DAY_SNAPSHOTS, THERAPIST_DAY_HISTORY, PATIENT_SESSION_HISTORY, DAY_HISTORY_EVENTS remain veryHidden.
- No unrelated worksheet protection changes.

Then manually verify in Excel:
1. Therapist statistics entry button opens the report.
2. Daily, weekly, monthly and overall report modes each refresh.
3. All four mode buttons and the previous-screen button can be clicked.
4. Previous returns to THERAPIST_DAILY.
5. Patient report entry and previous navigation still work.
6. A controlled report-render error after Unprotect leaves the report sheet protected, if the existing safe test harness supports it.

If preparation or audit fails, freeze BH. Never bypass the protection/button assertion, never switch DrawingObjects protection on merely to silence the check, and never proceed to broader protection.

## 5. Important COM semantics / current hypothesis

Excel Worksheet.Protect positional parameters are: Password, DrawingObjects, Contents, Scenarios.

Therefore Protect("", False, True, True) requests:
- no password,
- drawing objects not protected,
- cell contents protected,
- scenarios protected.

The previous failure only reported the generic message Therapist-history report buttons must remain clickable; it did not expose the actual property values. The new code includes both observed booleans in the error. **Do not claim the root cause is known until a new native run supplies those values.**

The installer also verifies shape names:
- ekaHistory_Daily
- ekaHistory_Weekly
- ekaHistory_Monthly
- ekaHistory_Overall
- ekaHistory_Back

Each must be visible and have a non-empty OnAction binding.

## 6. Non-negotiable safety rules

- Close Excel before pulling/running source tests or creating a fresh native smoke.
- Use python -m pytest, not bare pytest.
- Never save production/staging workbooks with openpyxl.
- Never edit the immutable baseline or trusted AZ/BF workbooks.
- Never overwrite an existing smoke destination. Each failed smoke is frozen and the next attempt gets a new suffix.
- Do not blanket-protect sheets until every legitimate human input range and every VBA/Python/form write path has been classified.
- THERAPIST_DAILY, MASTER_SCHEDULE, DAILY_INPUT, patient registries, settings and hidden storage have different write contracts. Do not infer writable ranges from UsedRange.Locked.
- The VBA scanner's 1,146 matches across 64 components are lexical candidates, not 1,146 proven writes.

## 7. Read order when a new conversation starts

1. Documentation/ONBOARDING.md — overall architecture and non-negotiable rules.
2. **This file** — exact current recovery state and next command.
3. Documentation/CURRENT_HANDOFF.md — latest detailed handoff.
4. Documentation/CHECKPOINT_H7_BF_MANUAL_ACCEPTANCE_AND_PROTECTION_NEXT_2026-11-04.md — accepted BF evidence and protection history.
5. Documentation/WORKSHEET_PROTECTION_WRITE_PATH_MAP_2026-11-04.md — per-sheet inventory and write-path risks.
6. Documentation/PROJECT_STATUS.md — project-wide status.

Do not use old chat history to override these repository records. If this file and a later committed checkpoint disagree, verify branch HEAD and prefer the newest explicit dated checkpoint.

## 8. Operator's immediate state

The operator has already received the PowerShell safety-stop after BG failure. No further local command should be run until the latest branch is pulled and the 15-test focused gate is green.

## Operator confirmation — focused test gate GREEN (2026-11-04)

The operator has now run the current focused suite and confirmed: **15 passed in 4.11s**. This supersedes the earlier statement in Section 3 that the latest 15-test result was pending. The source test gate is GREEN and the next native smoke is now authorized.

Next action: create a fresh `BH` from trusted `AZ` using the command in Section 4. Excel must be closed. Confirm the destination does not already exist. Do not use BG or BF as a source/destination. If BH preparation fails for any reason, freeze BH and stop. If it succeeds, run the read-only protection audit on BH and return its complete output before manual interaction.

## 9. Latest native result — BH frozen due patient-report protection regression (2026-11-04)

Operator supplied the complete read-only audit for:
`Excel/previews/UNIFIED_PATIENT_HISTORY_H7_20261104_BH.xlsm`.

Audit facts:
- ReadOnly=True; 23 worksheets; ProtectStructure=False; ProtectWindows=False.
- Workbook unchanged=True; audit completed.
- `THERAPIST_HISTORY_REPORT`: ProtectContents=True, ProtectDrawingObjects=False. This resolves the narrow BG failure at the persisted protection-state level; button visibility/OnAction was also checked by the installer, but native manual clicks have not yet been tested.
- Four technical history sheets remain protected and Visible=2 (veryHidden).
- **Regression:** `PATIENT_HISTORY_REPORT` reports ProtectContents=False and ProtectDrawingObjects=False. BF's accepted audit baseline had PATIENT_HISTORY_REPORT protected.
- Therefore BH fails the expected narrow-delta acceptance. Do not manually use, repair, or overwrite BH. Freeze it as failed evidence. Do not touch BF or AZ.

### Proven source-level cause

The regression is in the patient report installer, not in the therapist report protection code:

1. `Python/rehab_excel/patient_history_report_vba.py`, function `install_patient_report_components`, calls `report.Unprotect()`.
2. It clears the report and installs/positions the navigation UI.
3. It returns without re-protecting the blank report sheet.
4. `Python/tools/install_patient_history_ui.py`, function `_audit`, checked report visibility and toolbar wiring but did not check `ProtectContents` or `ProtectDrawingObjects`.

The patient report VBA renderer does protect the sheet after rendering and in its cleanup path. That does not protect the blank report immediately after a fresh install. BF's earlier manual report use caused the renderer to re-protect the sheet, masking this installer gap.

### Source correction committed after BH audit

- `81a8615`: added two regression tests in `Tests/test_patient_history_ui_h7.py` requiring installer-time protection and persisted-audit verification.
- `0fd7435`: `install_patient_report_components` now calls `report.Protect("", False, True, True)` after UI setup and verifies `ProtectContents=True`, `ProtectDrawingObjects=False`.
- `61ebf11`: `install_patient_history_ui._audit` now verifies the patient report protection state on both the writable pre-save audit and the separate read-only persisted reopen.

These are source changes only. **The tests have not yet been run by the operator.** Do not create another workbook until the combined source gate is green.

### Next source gate

With Excel closed, pull the branch and run:

```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }

python -m pytest Tests/test_therapist_history_ui.py Tests/test_patient_history_ui_h7.py Tests/test_h7_integrated_toolbar_and_patient_order.py Tests/test_h7_history_previous_screen_navigation.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: H7 report protection regression tests failed" }
```

Expected: **32 passed** (15 therapist UI + 6 patient UI + 8 integrated toolbar/patient-order + 3 previous-screen navigation tests). If the count differs or any test fails, stop and diagnose; do not prepare a workbook.

### Next native smoke after green source gate

Use trusted AZ as the source and a new unique destination **BI**, never BH/BG/BF:

```powershell
$source = ".\\Excel\\previews\\UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm"
$output = ".\\Excel\\previews\\UNIFIED_PATIENT_HISTORY_H7_20261104_BI.xlsm"

if (!(Test-Path -LiteralPath $source)) { throw "SAFETY STOP: AZ source missing" }
if (Test-Path -LiteralPath $output) { throw "SAFETY STOP: BI already exists; do not overwrite" }

python .\\Python\\tools\\prepare_h7_integrated_ui_smoke.py --source $source --output $output
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BI preparation failed; freeze BI" }

python .\\Python\\tools\\audit_workbook_protection.py $output
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BI read-only audit failed" }
```

Expected audit:
- `THERAPIST_HISTORY_REPORT`: ProtectContents=True; ProtectDrawingObjects=False.
- `PATIENT_HISTORY_REPORT`: ProtectContents=True; ProtectDrawingObjects=False.
- All five report sheets/buttons and navigation remain visible/bound as applicable.
- Four technical history sheets remain protected and veryHidden.
- All other worksheet protection states match the AZ/BF intended baseline; no unrelated changes.
- Workbook unchanged=True.

Only after this complete read-only audit passes may native manual clicks begin. Verify both report entry points, all four modes for each report, all mode buttons, previous-screen navigation, RefreshAll and existing application toolbar. Do not begin broad protection of the other sheets until this integrated H7 slice is accepted and the full write-path map is reconciled.

### Current artifact ledger
- AZ: trusted source, frozen, untouched.
- BF: accepted green native evidence, frozen, untouched.
- BG: failed therapist-report protection smoke, frozen.
- BH: therapist report protection now correct, but patient report protection regressed; frozen failed evidence.
- BI: next candidate only after the 33-test source gate passes.

