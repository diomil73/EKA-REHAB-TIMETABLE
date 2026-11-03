# H5 patient-session history checkpoint | 2026-11-03

## Status
H5 GREEN: native Excel proof confirmed by user outputs for service date 2026-10-30. Do not infer completion from source tests alone; evidence below is from Windows Excel.

## Immutable smoke artifacts
- `Excel/previews/UNIFIED_PATIENT_HISTORY_H5_20261103_AX.xlsm`: GREEN/FROZEN v1 close-day proof. Never modify or repair.
- `Excel/previews/UNIFIED_PATIENT_HISTORY_H5_20261103_AY.xlsm`: GREEN/FROZEN v1-to-v2 amendment proof. Never modify or repair.
- Prior unsuccessful smokes AR–AW: failed/frozen; never recycle.

## AX evidence: native first close
```text
CLOSE DAY NATIVE OUTCOME OK
SERVICE_DATE: 2026-10-30
SNAPSHOT: finalized:2026-10-30:v1 version=1 status=CURRENT_FINAL
COUNTS: planned=80 final=80 patient_absences=0 cancellations=0 no_shows=0 replacements=0 treatment_suspensions=0 approved_exceptions=0
THERAPIST_HISTORY_ROWS: 16
PATIENT_HISTORY_ROWS: 80
DAY_FINALIZED_EVENT: day-finalized:2026-10-30:v1
SHEET_STATE FINALIZED_DAY_SNAPSHOTS: veryHidden
SHEET_STATE THERAPIST_DAY_HISTORY: veryHidden
SHEET_STATE PATIENT_SESSION_HISTORY: veryHidden
SHEET_STATE DAY_HISTORY_EVENTS: veryHidden
```
User also confirmed native `Κλείσιμο ημέρας` success.

## AY evidence: native v1→v2
AY was hash-verified copy of AX (not a rebuilt blank smoke). The initial native close-day inspector reconfirmed v1 state; user applied a real change in the operational UI and saved the day, then reported:
```text
DAY AMENDMENT OUTCOME OK
Date: 2026-10-30
Version: 1 -> 2
Current snapshot: finalized:2026-10-30:v2
Therapist/day history rows: 16
Previous patient/session history rows: 80
Current patient/session history rows: 80
Operational CURRENT_FINAL patient rows: 80
```
This confirms v1 retained, v2 CURRENT_FINAL and operational reader scoped to current 80 rows. User described returning to MASTER after `Εφαρμογή & Συνέχεια`, then navigating back to the absences screen for the save action. Do not infer exact control clicked merely from informal wording; inspector proves amendment was persisted.

## Root cause and fix
AW diagnostic caught `OriginalTime expected=[08:30:00] actual=[0.3541666666666667] type=Double`. Excel `Value2` returns a numeric time serial; semantic readback expected Date. The numeric time-serial comparator fix was installed in both close-day and day-save modules; AX installer output confirmed `Persisted VBA reopen verification: OK`. Earlier commits: d6fb27b and 9ab6dde.

## UX follow-ups
1. Excel red spellcheck underline on Greek word `διακομισμένος`: user accepts and explicitly declines spending time now. Low priority, no action.
2. After application of absence/cancellation via `Εφαρμογή & Συνέχεια`, navigation returns to MASTER, forcing another visit to `Απουσίες / Ακυρώσεις` to reach day-save/close control. Record for future UX improvement, without disrupting passed H5 logic.
3. User previously expressed preference for one contextual `Αποθήκευση ημέρας` control rather than competing first-close and amendment controls. Defer design/implementation to a separate scoped change with regression and fresh smoke; don't alter AX or AY.

## Safety and resume instructions
- Production workbook immutable.
- AX and AY GREEN/FROZEN: read-only inspection allowed, no edits or rebuilds.
- Same-session day-save proof completed; no need to repeat H5.
- Any future UX change needs source tests first and a fresh uniquely named native workbook; never retry failed/frozen artifacts.
- Preserve authoritative snapshot/version/event history and protection/rollback semantics.
