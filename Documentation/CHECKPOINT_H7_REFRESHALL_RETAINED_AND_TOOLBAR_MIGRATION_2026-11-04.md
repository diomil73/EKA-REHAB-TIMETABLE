# H7 RefreshAll retained and integrated toolbar migration staged | 2026-11-04

## Operator evidence: full RefreshAll macro inspected

Read-only audit of frozen diagnostic BE:
```text
FULL RefreshAll PROCEDURE Module3:1-19
Public Sub RefreshAll()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Call RefreshTherapistDaily
    Call RefreshTherapistAttendanceAndConflicts
    Call BuildReplacementSuggestions_V4
    Call ColorInfectiousInMaster
    Call ColorReplacementMarks
    Call RestoreReturnedTherapists
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Όλα ενημερώθηκαν.", vbInformation
End Sub
```
There are no workbook connections, pivot caches, query tables or pivot tables, but the macro calls six operational procedures. Therefore the Refresh button is **not obsolete**. Preserve macro and functionality. Do not remove the action.

The legacy shape is a Form Control named `Button 1`, with empty caption, bound to workbook-qualified `RefreshAll`, at far right of THERAPIST_DAILY. The problem is its legacy geometry/appearance only.

## Source change: migrate control safely in a fresh workbook

- `Python/tools/install_patient_history_ui.py` commits `7d5af44`, `54e3b2e`, `76011e7`:
  - During the final H7 installer stage, find exact existing THERAPIST_DAILY shapes whose OnAction resolves to `RefreshAll`.
  - Require exactly one legacy action if canonical `ekaToolbar_RefreshAll` is absent. Fail closed for zero or multiple matches.
  - Create a new styled AutoShape named `ekaToolbar_RefreshAll`, caption `ΑΝΑΝΕΩΣΗ`, width 92, height 26, black bold font, same blue-grey fill, bound to the unchanged `RefreshAll` macro.
  - Verify new binding before deleting the one exact legacy form control.
  - Reposition the canonical RefreshAll immediately after Save & Exit on the same toolbar row.
  - Place therapist stats immediately after RefreshAll and patient stats after therapist stats.
  - Persisted reopen audit requires Module3.RefreshAll, one and only one canonical RefreshAll shape, correct caption, binding, visibility and nonoverlap.
- `Python/tools/inspect_h7_integrated_ui.py` commits `edae5b0`, `db8c9ef`, `b08a8b4`: independent read-only audit checks same toolbar arrangement, caption, unique RefreshAll action, and presence of Module3 procedure.
- `Tests/test_h7_integrated_toolbar_and_patient_order.py` commits `ba50078`, `0036a54`, `8673c4c`: source assertions for migration, preserved macro, ordered buttons and independent audit.
- None of these new source changes has been operator-tested yet.
- The previous-screen navigation changes are still staged in patient and therapist report VBA. Operator has not supplied exact pytest summary for the prior 18-test command, so re-run relevant suite below.

## Expected toolbar order

On THERAPIST_DAILY, same top coordinate as Save & Exit:
1. `← MASTER`
2. `Save & Exit`
3. `ΑΝΑΝΕΩΣΗ`
4. `Στατιστικά θεραπευτών`
5. `Στατιστικά ασθενών`

Approximate geometry from BE original: Back x=4,w=82; Exit x=92,w=70. New expected: Refresh x=170,w=92; Therapist x=270,w=146; Patient x=424,w=146. All share y=563.6, based on current BE's main toolbar row. Native installer derives positions from Exit shape rather than hardcoded y.

## Next source gate (Excel closed)

```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
python -m pytest Tests/test_h7_integrated_toolbar_and_patient_order.py Tests/test_h7_history_previous_screen_navigation.py Tests/test_therapist_history_ui.py Tests/test_patient_history_ui_h7.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: H7 toolbar/navigation tests failed" }
```

Expected **26 passed** (8+3+10+5), pending operator proof.

## Next native candidate

Only after 26 tests GREEN, create a unique new BF from frozen AZ with ordered H6+H4+H7 integrated preparation. Do not modify or overwrite BE. The new H7 installer migrates the legacy Button 1 and preserves macro action. Then run independent read-only integrated audit. If any failure, freeze BF DIAGNOSTIC; do not repair/reinstall.

Native acceptance:
- Verify exact toolbar order and no detached RefreshAll shape.
- Click RefreshAll once in BF and verify its six operational effects / success message, while observing no data loss or duplicate writes. If any side effect unclear, stop and inspect before repeated click.
- Test both statistics entry buttons, daily/weekly/monthly/overall and "← Προηγούμενη" direct return to THERAPIST_DAILY.
- Verify registry 93 overlay, 16 zero rows and preserved 80 scheduled / 77 performed / 3 absences.
- Final worksheet protection gate remains outstanding. BE and previous smokes remain untouched.

## Operator verification update

Operator ran the complete source gate and reported: `26 passed in 4.87s`.
This confirms all 8 integrated toolbar tests, 3 previous-screen navigation tests, 10 therapist-history UI tests, and 5 patient-history UI tests passed on the current branch.

## Immediate next operator step: create BF

Excel must be fully closed. From repository root, run:
```powershell
python .\Python\tools\prepare_h7_integrated_ui_smoke.py `
  --source ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm" `
  --output ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H7_20261104_BF.xlsm"
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BF preparation failed; freeze BF and do not retry/repair" }
```

The script SHA-256 verifies a fresh copy and runs, in order:
1. `install_unified_day_save_ui.py`
2. `install_therapist_history_ui.py`
3. `install_patient_history_ui.py` (includes the RefreshAll migration)

Expected installer evidence:
- H6 day-save install succeeds and persisted reopen verifies.
- H4 therapist history install succeeds and persisted reopen verifies.
- H7 patient history install succeeds; RefreshAll is preserved, styled and placed on the main toolbar; both statistics buttons align after it; persisted reopen verifies.
- Final message: `H7 INTEGRATED UI INSTALL OK`.

**Critical stop rule:** If any installer fails, preserve BF as diagnostic and do not rerun an installer or overwrite BF. Report the full output.

## After successful preparation

Run independent read-only audit:
```powershell
python .\Python\tools\inspect_h7_integrated_ui.py `
  ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H7_20261104_BF.xlsm"
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BF integrated audit failed" }
```
Expected: `H7 INTEGRATED AUDIT OK; workbook not changed`, with toolbar order Back, Save & Exit, RefreshAll, therapist statistics, patient statistics. The audit also verifies Module3.RefreshAll and one canonical visible refresh control.

Do not click RefreshAll until the read-only audit passes. Then test it once and inspect all six effects from Module3: RefreshTherapistDaily, RefreshTherapistAttendanceAndConflicts, BuildReplacementSuggestions_V4, ColorInfectiousInMaster, ColorReplacementMarks, RestoreReturnedTherapists. Verify no duplicate writes or unintended changes before further interaction.

AZ and BE remain untouched. BF is a disposable smoke only, not production.

## Operator native BF audit result

Operator ran the independent read-only audit on:
`Excel/previews/UNIFIED_PATIENT_HISTORY_H7_20261104_BF.xlsm`

Exact reported outcome:
- ReadOnly=True; `H7 INTEGRATED AUDIT OK; workbook not changed`.
- MASTER SaveDay macro remains `SaveOperationalDay`, at left=434, top=11; ExitRight=426.
- THERAPIST_DAILY RefreshAll macro binding is `RefreshAll`, left=170, top=563.6, width=92; caption exactly `ΑΝΑΝΕΩΣΗ`.
- Therapist report macro `OpenTherapistHistoryReport`, left=270, top=563.6.
- Patient report macro `OpenPatientHistoryReport`, left=424, top=563.6.
- VBA components Module3, modUnifiedDaySave, modTherapistHistoryReport, modPatientHistoryReport all present.
- Both report sheets visible; all technical history sheets veryHidden.
- No audit failure.

**BF integrated installation and read-only audit are GREEN.** This is still a disposable smoke workbook, not production and not a release freeze.

## Next native manual test sequence (BF only)

Open `UNIFIED_PATIENT_HISTORY_H7_20261104_BF.xlsm` in Excel. Do not use production workbook.

1. Visually confirm same-row toolbar order: `← MASTER`, `Save & Exit`, `ΑΝΑΝΕΩΣΗ`, `Στατιστικά θεραπευτών`, `Στατιστικά ασθενών`.
2. Click patient statistics. Confirm report opens, review all four modes (daily, weekly, monthly, overall), and confirm the new `← Προηγούμενη` button returns directly to THERAPIST_DAILY.
3. Click therapist statistics. Confirm report opens and `← Προηγούμενη` returns directly to THERAPIST_DAILY.
4. Only after report navigation succeeds, click `ΑΝΑΝΕΩΣΗ` once. Expect the original `RefreshAll` six calls and completion message `Όλα ενημερώθηκαν.`. Record any error or unexpected change; do not repeatedly click if any effect is unclear.
5. Verify the daily therapist sheet remains usable and the two statistics buttons remain visible/aligned after refresh.

Do not click `Save & Exit` or `SaveOperationalDay` as part of this H7 toolbar test. Do not finalize/amend a day. The refresh test is isolated from day finalization.

If any error occurs, close BF without saving if safe, preserve the file as diagnostic, and report exact message/screenshot. Do not rerun installers on BF.

After manual pass, next gates: rerun read-only integrated audit after Excel is closed; then focused historical coverage/report-mode checks; final protection/editability audit remains a release gate.
