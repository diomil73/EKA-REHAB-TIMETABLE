# H7 registry overlay 19-test gate GREEN; BE native pending | 2026-11-04

## Confirmed test evidence
Operator supplied:
```text
19 passed in 3.46s
```
Source gate executed for:
- `Tests/test_h7_registry_zero_overlay_and_visibility.py` (9)
- `Tests/test_patient_history_reports_h7.py` (7)
- `Tests/test_h7_patient_mode_switch_and_toolbar_fix.py` (3)

These validate current-registry zero-activity overlay, previous history-only compatibility, row-height and UX source fixes. **No claim of native BE success.**

## Prior read-only facts
- Frozen AZ/diagnostic BD: October 30, 2026, 80 CURRENT_FINAL sessions, 77 unique active patient IDs, 77 performed, 3 patient absences.
- Registry: 93 patients. Missing activity on that finalized date for 16 registry IDs (02, 06, 15, 19, 20, 21, 24, 25, 30, 39, 41, 50, 64, 69, 73, 75).
- 18, 22 present in saved BD monthly report, visible and unfiltered (rows 20 and 21, prior height 13.8).
- BD runtime 'Overall' once failed then worked; native root cause unconfirmed. BD not accepted; do not change.
- Integrated BD persisted structure audited GREEN but UX fail remains; AZ/BB frozen GREEN.

## NEXT: fresh BE native installation (Excel fully closed)

The orchestrator `Python/tools/prepare_h7_integrated_ui_smoke.py` copies AZ to a **new nonexisting** BE, byte-verifies SHA256, installs in order: H6 MASTER-only save layout, H4 therapist statistics, H7 patient statistics. Each installer persists, reopens and audits. Source changes include zero-registry patient rows, more legible height, unmerge-before-clear, error detail, matching black font and navigation layout.

```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: pull failed" }
$be = ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H7_20261104_BE.xlsm"
if (Test-Path $be) { throw "SAFETY STOP: BE already exists; do not overwrite or reinstall" }
python .\Python\tools\prepare_h7_integrated_ui_smoke.py --output $be
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BE integrated installer failed; freeze BE" }
python .\Python\tools\inspect_h7_integrated_ui.py $be
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BE persisted structural audit failed; freeze BE" }
```

If BE is already present, do not rerun installer, inspect only read-only and capture provenance. If either fails, freeze BE FAILED/DIAGNOSTIC; next unique suffix.

## BE native acceptance, only after installer + audit GREEN
- MASTER Save Day next to Save & Exit, not overlapping.
- DAILY_INPUT no legacy save buttons or stats entry.
- THERAPIST_DAILY two aligned readable statistics buttons.
- Patient daily, weekly, monthly, overall, daily again and back, no generic failures; inspect rows for 15,18–22 and 01/02/... numeric order.
- **Current registry** overlay displays 93 patients on the one finalized day, 16 zero-activity and historical sums 80 scheduled, 77 performed, 3 absences. Zero rows are not recorded sessions and must be labeled current-registry overlay.
- Therapist statistics still accessible.
- On success freeze BE GREEN and checkpoint; on failure freeze BE FAILED and diagnose, do not reinstall.
- No re-run of previously successful H5/H6 v1/v2 day-save proofs. Production immutable. Final per-cell edit-protection release blocker remains pending.
