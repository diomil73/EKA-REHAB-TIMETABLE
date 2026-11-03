# H7 BE screenshot: unexpected Refresh button and toolbar placement | 2026-11-04

## Native user evidence

Operator shared an Excel screenshot of `UNIFIED_PATIENT_HISTORY_H7_20261104_BE.xlsm` on `THERAPIST_DAILY`. Visible issues:
- A conspicuous button labeled `ΑΝΑΝΕΩΣΗ` at the far right of the schedule, detached from the main navigation.
- `Στατιστικά θεραπευτών` and `Στατιστικά ασθενών` placed beneath the restoration controls instead of the main toolbar row; geometry and spacing are not acceptable.
- `← MASTER` and `Save & Exit` are visibly paired at lower left on the main toolbar row. The two statistics buttons should be aligned to that row and kept clear of restoration controls and table.

Screenshot itself is **not evidence** that Refresh is safe to delete or that a full functional H7 test passed. No assertion of BE GREEN/FROZEN. Do not mutate BE or any previous smoke.

## Diagnostics committed

`Python/tools/audit_h7_therapist_daily_shapes.py` commit `bcd60a7` enumerates all `THERAPIST_DAILY` shapes read-only, with caption/name/OnAction/position/dimensions/font color. This will determine which macro (if any) is behind the anomalous `ΑΝΑΝΕΩΣΗ` button and confirm exact toolbar coordinates.

With Excel closed:
```powershell
git pull --ff-only origin work/history-statistics-rebuild
python .\Python\tools\audit_h7_therapist_daily_shapes.py ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H7_20261104_BE.xlsm"
```
Use this output to determine whether refresh is obsolete or essential before deleting any shape. If refresh is a legacy unused control, remove it in next installer source and add regression + persisted audit; do not remove blindly.

## Proposed follow-up

Align H4/H7 entry buttons with the existing `ekaToolbar_ExitApp` on `THERAPIST_DAILY` (same top row, sequential to Save & Exit), not with `UsedRange` last row. Keep black text and nonoverlapping geometry. Verify no overlap with cell labels/restore actions in the screenshot. Source changes and tests first, then a **new uniquely suffixed BF** candidate from AZ, not a repair of BE. Ensure release-related sheet protection remains pending.
