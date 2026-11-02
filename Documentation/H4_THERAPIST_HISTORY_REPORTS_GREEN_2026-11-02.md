# H4 Therapist Historical Reports — Green Proof

Date: 2026-11-02
Branch: `work/history-statistics-rebuild`

## Status

H4 reporting backend is green for daily, weekly, monthly, and overall therapist history views.

## Canonical proof workbook

`Excel/previews/UNIFIED_DAY_SAVE_UI_20261102_AM.xlsm`

AM remains frozen green evidence from H3 and is used read-only for H4 report verification.

## Source gate

Operator-confirmed:

```text
17 passed in 1.25s
```

Focused tests:

- `Tests/test_therapist_day_history.py`
- `Tests/test_therapist_history_reports.py`

## Daily report proof

The report reads finalized historical facts only and excludes superseded snapshots from active reporting.

Observed:

```text
THERAPIST HISTORY REPORT OK
Rows: 16
```

The amended date 2026-10-30 is reported from v2 only. v1 remains preserved in history but is not double-counted.

## Weekly report proof

Observed:

```text
THERAPIST WEEKLY HISTORY REPORT OK
Rows: 16
```

For the current one-day dataset, the aggregation window is 2026-10-26 through 2026-11-01 with `days=1` for each therapist.

## Monthly report proof

Observed:

```text
THERAPIST MONTHLY HISTORY REPORT OK
Rows: 16
```

The report correctly aggregates the available finalized data into the 2026-10-01 through 2026-10-31 window.

## Overall report proof

Observed:

```text
THERAPIST OVERALL HISTORY REPORT OK
Rows: 16
```

With one finalized service date currently available in AM, the overall date range is 2026-10-30 through 2026-10-30.

## Reporting facts proven

The report surfaces, per therapist and period:

- reported days,
- planned occupied timeslots,
- actual occupied timeslots,
- planned sessions,
- active sessions,
- replacements received,
- released load from cancellations/no-shows,
- treatment suspensions affecting load,
- planned capacity utilization,
- actual capacity utilization.

The history model preserves the distinction between occupied timeslots and session count, so shared-session scenarios are not flattened.

Examples from the AM proof include:

- `Σάρρας`: 6 occupied timeslots / 7 sessions,
- `Χρήστου`: 5 occupied timeslots / 6 sessions,
- `Σκαρώνης`: 7 occupied timeslots with 116.7% capacity utilization,
- `Πέτσιος`: zero reported load for the finalized day.

## Safety contract

- Reports are read-only.
- Only `CURRENT_FINAL` snapshots are reportable.
- `SUPERSEDED` snapshots remain preserved but are excluded from current totals.
- Multiple `CURRENT_FINAL` snapshots for the same service date fail closed.
- Any current snapshot with unresolved sessions fails closed.
- No historical worksheet is rewritten by reporting.

## H4 remaining product work

The reporting backend is complete for the first H4 scope. The remaining product-facing step is to expose therapist history/statistics inside the Excel application so the operator does not need PowerShell/CLI access.

Do not alter the frozen H3/H4 proof workbook in place. Any UI smoke must use a fresh uniquely named workbook derived from an approved green source.
