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

`Excel/previews/DAILY_CHANGE_LEDGER_V83.xlsm` is the latest native checkpoint.

Stress case:

- date `2026-10-19`,
- patient `ΜΕΝΤΗΣ`,
- original `Σκαρώνης 09:15`,
- selected replacement `Μηλιδάκης 10:45`,
- candidate load `load 5/6 | patients 6`,
- source unchanged,
- VBA preserved,
- ledger installed.

Native behavior now confirmed:

- `THERAPIST_DAILY` shows the effective replacement and keeps the source-position trace,
- `MASTER` shows the same effective daily assignment while the move is active,
- individual restore returns both views to `Σκαρώνης 09:15`.

Therefore MASTER daily-state integration is **native green for the V83 stress case**.

## Current display fix

V83 still shows `####` in the left `THERAPIST_DAILY` time column.

Root cause identified: ledger refresh reset column A to width 4.

Implemented fix:

- column A width 8,
- `A2:A8,A12:A18` forced to `hh:mm`,
- centered time axis,
- regression preventing width 4 from returning.

Status: **implemented, focused/native verification pending**.

## Remaining native cleanup backlog

- verify the time-axis fix removes all `####`,
- orange text remains reserved for robotic participation only,
- moved infectious ΜΕΝΤΗΣ must not inherit robotic-orange styling,
- remove obsolete legacy `ΑΝΑΝΕΩΣΗ` button,
- keep the useful source-slot trace showing the moved patient and arrow.

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

1. Run focused ledger regression after the time-axis fix.
2. Build fresh preview and confirm all left-side times show as `hh:mm` with no `####`.
3. Fix infectious vs robotic styling for moved patients.
4. Remove legacy `ΑΝΑΝΕΩΣΗ` button.
5. Implement/verify orphan final outcomes and statistics.
6. Test selected/all restore with multiple changes.
7. Replace preview snapshot restore with authoritative overlay removal + full daily recomputation.
8. Resume outpatient recurring create/edit native gate.
9. Continue remaining native checks and branch/PR hardening.
10. Clean up Python 3.14 ZipFile teardown warnings when convenient.

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
