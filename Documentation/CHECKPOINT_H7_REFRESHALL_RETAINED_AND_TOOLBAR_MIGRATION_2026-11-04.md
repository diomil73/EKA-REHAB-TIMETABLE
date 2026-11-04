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
