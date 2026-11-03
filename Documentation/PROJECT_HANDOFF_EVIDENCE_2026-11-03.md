# Project handoff: evidence and historical diagnosis

**Repository:** `diomil73/EKA-REHAB-TIMETABLE`  
**Branch:** `work/history-statistics-rebuild`  
**Verified at handoff:** 2026-11-03. All cited operator outputs were supplied in project conversation; source/native claims are distinguished.

## Chronology and proof

| Stage | Status | Immutable artifact | Evidence |
|---|---|---|---|
| H2 close day | GREEN/FROZEN | `Excel/previews/UNIFIED_CLOSE_DAY_HISTORY_20261101_Z.xlsm` | v1 finalized, therapist and event history |
| H3 same-session day save | GREEN/FROZEN | `Excel/previews/UNIFIED_DAY_SAVE_UI_20261102_AM.xlsm` | v1→v2 authoritative save without Excel reopen |
| H4 therapist history UI | GREEN/FROZEN | `Excel/previews/UNIFIED_THERAPIST_HISTORY_UI_20261102_AQ.xlsm` | Operator switched daily/weekly/monthly/overall repeatedly; buttons remained usable |
| H5 initial close | GREEN/FROZEN | `Excel/previews/UNIFIED_PATIENT_HISTORY_H5_20261103_AX.xlsm` | Native close outcome OK, 80 patient rows, 16 therapist rows |
| H5 amendment | GREEN/FROZEN | `Excel/previews/UNIFIED_PATIENT_HISTORY_H5_20261103_AY.xlsm` | Native amendment outcome OK, v1→v2, previous/current 80 patient rows |
| H6 functional UX | GREEN/FROZEN | `Excel/previews/UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm` | First close successful, native amendment outcome v1→v2, CURRENT_FINAL 80 rows |
| H6 final visual layout | GREEN/FROZEN | `Excel/previews/UNIFIED_PATIENT_HISTORY_H6_20261103_BB.xlsm` | Persisted read-only audit and user visual confirmation; no day-save rerun on BB |

**FAILED/FROZEN:** H5 AR–AW, H6 BA. These are diagnostic/historical evidence only; **never attempt to repair or overwrite**. H5 AW exposed the root cause; H6 BA exposed missing unified UI installation.

## H5 AX inspector output (material fields)

```text
CLOSE DAY NATIVE OUTCOME OK
SERVICE_DATE: 2026-10-30
SNAPSHOT: id=finalized:2026-10-30:v1 version=1 status=CURRENT_FINAL
COUNTS: planned=80 final=80 patient_absences=0 cancellations=0
THERAPIST_HISTORY_ROWS: 16
PATIENT_HISTORY_ROWS: 80
DAY_FINALIZED_EVENT: id=day-finalized:2026-10-30:v1
SHEET_STATE FINALIZED_DAY_SNAPSHOTS: veryHidden
SHEET_STATE THERAPIST_DAY_HISTORY: veryHidden
SHEET_STATE PATIENT_SESSION_HISTORY: veryHidden
SHEET_STATE DAY_HISTORY_EVENTS: veryHidden
```

## H5 AY and H6 AZ native amendment inspector

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

AX and AY are separate archival proofs. AZ uses the contextual save dispatcher and passed functional amendment. For AZ v1 there was operator UI confirmation but no separately pasted v1 inspector result; do not invent one.

## Excel serial-time failure and resolution

AW diagnostic: `OriginalTime expected=[08:30:00] actual=[0.3541666666666667] type=Double`. Excel `Value2` stores time as fraction of a day. Original VBA readback comparator handled Date but not numeric Excel time serial. Fixes covered close-day and amendment paths; relevant commits `d6fb27b`, `9ab6dde`. AX installer confirmed persisted VBA reopen verification before successful native close.

## H6 BA failure and BB correction

BA read-only audit reported MASTER only `ekaToolbar_ExitApp`; DAILY_INPUT still had `ekaToolbar_CloseDay` and `ekaToolbar_DayAmendment`; `DISPATCHER: MISSING (com_error)`. Do not assume the cause was an installer bug: the persisted artifact simply lacked the final installer result. BA frozen.

BB audit:
```text
SHEET MASTER_SCHEDULE:
  ekaToolbar_ExitApp: OnAction=ExitApplication; Left=348.0; Top=11.0; Width=78.0; Visible=True
  ekaToolbar_SaveOperationalDayMaster: OnAction=SaveOperationalDay; Left=434.0; Top=11.0; Width=154.0; Visible=True
SHEET DAILY_INPUT:
  ekaToolbar_ExitApp: OnAction=ExitApplication; Left=216.0; Top=621.2; Width=70.0; Visible=True
DISPATCHER: PRESENT
```
User confirmed BA was wrong and BB layout "όλα οκ". Final desired UX: `Απουσίες / Ακυρώσεις` -> `Εφαρμογή & Συνέχεια` -> MASTER -> **only visible `Αποθήκευση ημέρας` control**. The dispatcher routes to first close or amendment via current snapshot state; ambiguous state stops safely.

## H7 actual proof and code

Tests:
- `Tests/test_patient_statistics_h7.py`: **4 passed / 2.25s**.
- `Tests/test_patient_statistics_h7.py` + `Tests/test_historical_statistics_h7.py`: initial 1 failed/10 passed because missing patient row triggered generic totals disagreement before specific missing-row error. Fix commit `9810182` reordered validations; retry **11 passed / 2.58s**.
- Above plus `Tests/test_period_statistics_h7.py`: **15 passed / 3.29s**.

Production-independent read-only monthly report on AZ:
```json
{
  "ok": true,
  "period": "month",
  "report_rows": [{
    "period": "2026-10",
    "start_date": "2026-10-01",
    "end_date": "2026-10-31",
    "finalized_days": 1,
    "scheduled": 80,
    "performed": 77,
    "patient_absences": 3,
    "cancellations": 0,
    "replacements": 0
  }],
  "workbook_unchanged": true
}
```
This supports read-only monthly statistics from existing v2 history, with 77+3=80. It does **not** prove a workbook-embedded statistics screen, multi-month coverage, or final sheet lockdown.

Relevant code:
- `Python/rehab_excel/patient_statistics.py` patient rollups, commit `af2ec40`.
- `Python/rehab_excel/historical_statistics.py` current snapshot reconciliation, commits `e467e7e`, `9810182`.
- `Python/rehab_excel/period_statistics.py` day/ISO-week/month rollups, commit `e550158`.
- `Python/tools/report_period_statistics.py` read-only CLI, commit `279687e`.
- `Tests/test_patient_statistics_h7.py`, `Tests/test_historical_statistics_h7.py`, `Tests/test_period_statistics_h7.py`.

## Supporting older checkpoints

- `Documentation/CHECKPOINT_H4_THERAPIST_HISTORY_2026-11-03.md`
- `Documentation/CHECKPOINT_H5_PATIENT_HISTORY_2026-11-03.md`
- `Documentation/CHECKPOINT_H6_CLOSED_H7_STARTED_2026-11-03.md`
- `Documentation/CHECKPOINT_H6_MASTER_ONLY_SAVE_2026-11-03.md`
- `Documentation/CHECKPOINT_H7_AGGREGATION_2026-11-03.md`
- `Documentation/BACKLOG_FINAL_WORKSHEET_PROTECTION_2026-11-03.md`

Older instructions may be stale after later green results; this dated evidence document is the chronological source of truth.
