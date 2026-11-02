# H4 Therapist History Reports Checkpoint — 2026-11-02

## Status

The first H4 historical therapist reporting slices are green against the frozen H3 proof workbook.

Frozen source artifact used read-only:

`Excel/previews/UNIFIED_DAY_SAVE_UI_20261102_AM.xlsm`

No workbook cells were written by the H4 report commands.

## Daily report proof

Command:

```powershell
python .\Python\tools\report_therapist_history.py `
  --workbook ".\Excel\previews\UNIFIED_DAY_SAVE_UI_20261102_AM.xlsm"
```

Observed operator result:

- `THERAPIST HISTORY REPORT OK`
- `Rows: 16`
- all rows came from snapshot version `v2`
- the superseded `v1` was not double-counted
- planned and actual occupied slots/sessions were reported
- replacements, released load, suspensions and capacity utilization were reported

Notable observed examples:

- Σκαρώνης: 7/6 occupied slots, 116.7% utilization
- Σάρρας: 6 occupied slots and 7 active sessions, preserving shared-slot semantics
- Χρήστου: 5 occupied slots and 6 active sessions, preserving shared-slot semantics
- Πέτσιος: zero planned and actual load for the tested day

## Weekly report proof

Command:

```powershell
python .\Python\tools\report_therapist_history.py `
  --workbook ".\Excel\previews\UNIFIED_DAY_SAVE_UI_20261102_AM.xlsm" `
  --weekly
```

Observed operator result:

- `THERAPIST WEEKLY HISTORY REPORT OK`
- `Rows: 16`
- period `2026-10-26..2026-11-01`
- `days=1` for this proof workbook, because only one finalized service date is present
- daily workload facts aggregate without losing shared-slot/session distinctions

## Report safety contract

- report inputs are the durable `FINALIZED_DAY_SNAPSHOTS` and `THERAPIST_DAY_HISTORY` stores
- only `CURRENT_FINAL` snapshots are reportable
- superseded versions remain preserved in history but are excluded from current operational reporting
- multiple `CURRENT_FINAL` snapshots for one service date fail closed
- current snapshots with unresolved sessions fail closed
- duplicate therapist/day history rows fail closed
- reporting is read-only

## Current H4 capability

Green:

1. daily therapist historical workload
2. weekly therapist historical workload aggregation
3. planned vs actual occupied timeslots
4. planned vs active session counts
5. replacements received
6. released load from cancellations/no-show
7. treatment suspensions affecting load
8. planned/actual capacity utilization when capacity is available

## Next H4 slice

Add monthly and all-history therapist aggregates/trends using the same `CURRENT_FINAL`-only contract. Do not create a second history store and do not modify the frozen AM proof artifact.
