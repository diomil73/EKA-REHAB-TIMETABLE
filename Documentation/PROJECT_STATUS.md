# Project Status

Last updated: 2026-10-17

## Current stage

The project remains on branch `feature/outpatient-scheduling`.

The patient-centric multi-treatment scheduling milestone is closed. The active operational milestone is **one concrete-date effective state across MASTER + THERAPIST_DAILY + ledger + statistics**, with outpatient recurring create/edit still open behind it.

Latest full-suite checkpoint after cleanup: **699 passed, 19 warnings**.
Latest focused operational regression checkpoint: **31 passed in 2.04s**.
Latest MASTER daily-overlay gate: **24 passed, 5 warnings in 4.39s**.

The 5 warnings are Python 3.14 `PytestUnraisableExceptionWarning` messages from `ZipFile.__del__` during workbook-test teardown. They are currently non-blocking technical debt, not a failed functional gate.

## Current native checkpoint

`Excel/previews/DAILY_CHANGE_LEDGER_V84.xlsm` is the latest native checkpoint.

Stress case:

- date `2026-10-19`,
- patient `ΜΕΝΤΗΣ`,
- original `Σκαρώνης 09:15`,
- selected replacement `Μηλιδάκης 10:45`,
- source unchanged,
- VBA preserved,
- ledger installed.

Native behavior confirmed:

- `THERAPIST_DAILY` shows the effective replacement and keeps the source-position trace,
- `MASTER` shows the same effective daily assignment while the move is active,
- individual restore returns both views to `Σκαρώνης 09:15`,
- left-side time axis now shows readable `hh:mm`; the `####` issue is closed.

Therefore MASTER daily-state integration, individual restore, and time-axis display are **native green for the V84 stress case**.

## Current source fixes awaiting native verification

Two presentation issues remain visible in V84:

1. moved non-robotic infectious ΜΕΝΤΗΣ appears orange at destination,
2. obsolete legacy `ΑΝΑΝΕΩΣΗ` button remains.

Implemented fixes:

- rich-text cell rewrite resets stale whole-cell font color to default black before semantic runs are applied,
- only actual `robotic_orange` text runs receive orange,
- ledger preview installer deletes only `THERAPIST_DAILY` shapes whose exact caption is `ΑΝΑΝΕΩΣΗ`, preserving ledger controls.

Regression coverage added for both fixes.

Status: **implemented, focused tests + native preview verification pending**.

## Orphan workflow locked

The daily `Ορφανοί` list is an operational queue. Every unresolved/displaced patient must end as either:

- placed in a valid therapist/time slot, or
- `Αναστολή θεραπείας` for that date only.

Suspension is a daily overlay/status, never a recurring delete, and feeds statistics as a non-performed treatment with explicit reason. No orphan may silently disappear.

## Daily operational rules locked

- capacity/load is occupied timeslots,
- shared slots may host multiple patients but consume one timeslot,
- cancellations release same-day capacity,
- therapist absences and cancellations may arrive in any order,
- recomputation is idempotent,
- displaced existing patients precede new patients,
- remaining orphans are explicitly placed or suspended,
- recurring/base schedules are never destructively edited for one-day facts,
- date-specific moves are reversible,
- mixed inpatient/outpatient presentation ambiguity must not block valid movement.

## Accepted MASTER UX

Native V81 remains green:

- four action buttons horizontally at the top,
- treatment/column headers visible,
- rows 1-4 frozen during vertical scrolling,
- no UserForm/sidebar/CommandBars workaround.

## Outpatient architecture already in place

The branch contains inpatient/outpatient types, separate authoritative `OUTPATIENT_SCHEDULE`, unified base reads, outpatient presentation, daily operational integration, cancellation/no-show overlays, and outpatient schedule registration/suggestion/conflict infrastructure.

Outpatient recurring create/edit native verification remains open after the current daily-state integration work.

## Immediate roadmap

1. Run focused tests for stale-font reset and legacy refresh-button removal.
2. Build fresh guided replacement preview.
3. Verify ΜΕΝΤΗΣ at `Μηλιδάκης 10:45` is not orange unless actually robotic.
4. Verify legacy `ΑΝΑΝΕΩΣΗ` is gone while ledger controls remain.
5. Verify individual restore still works and MASTER/THERAPIST_DAILY stay synchronized.
6. Implement/verify orphan final outcomes and statistics.
7. Test selected/all restore with multiple changes.
8. Replace preview snapshot restore with authoritative overlay removal + full daily recomputation.
9. Resume outpatient recurring create/edit native gate.
10. Continue remaining native checks and branch/PR hardening.
11. Clean up Python 3.14 ZipFile teardown warnings when convenient.

## Safety constraints

- Never write directly to the baseline `.xlsm` during development or smoke testing.
- Preview workflows operate on copied `.xlsm` files only.
- Preserve/verify VBA when rebuilding previews.
- Keep current scheduling behavior backward-compatible unless an explicit business rule changes it.
- Windows test command is always `python -m pytest`.

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests
