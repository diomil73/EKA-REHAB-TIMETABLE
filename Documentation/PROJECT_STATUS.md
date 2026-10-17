# Project Status

Last updated: 2026-10-17

## Current stage

The project remains on branch `feature/outpatient-scheduling`.

The patient-centric multi-treatment scheduling milestone is closed. The active operational milestone is now **one concrete-date effective state across MASTER + THERAPIST_DAILY + ledger + statistics**, with outpatient recurring create/edit still open behind it.

Latest full-suite checkpoint after cleanup: **699 passed, 19 warnings**.
Latest focused operational regression checkpoint: **31 passed in 2.04s**.

## Current native checkpoint

`Excel/previews/DAILY_CHANGE_LEDGER_V82.xlsm` builds successfully and the individual restore flow is native green.

Stress case:

- date `2026-10-19`,
- patient `ΜΕΝΤΗΣ`,
- original `Σκαρώνης 09:15`,
- selected replacement `Μηλιδάκης 10:45`,
- candidate load `load 5/6 | patients 6`,
- source unchanged,
- VBA preserved,
- ledger installed,
- per-row restore returns the patient correctly to the original position.

## Confirmed integration bug

While the daily move is active, `THERAPIST_DAILY` shows the effective replacement correctly, but the timetable on `MASTER` remains on the recurring/original assignment.

This is not accepted behavior.

Required invariant:

`DailySessionState -> THERAPIST_DAILY + MASTER timetable + ledger + statistics`

MASTER must show the effective therapist/time while an overlay is active and return to the original therapist/time after restore.

## V82 native cleanup backlog

- left time column in `THERAPIST_DAILY` shows `###` in some rows; widen/fix `hh:mm`,
- orange text remains reserved for robotic participation only,
- moved infectious patient ΜΕΝΤΗΣ must not inherit robotic-orange styling,
- remove obsolete legacy `ΑΝΑΝΕΩΣΗ` button on the right side,
- keep the useful source-slot trace showing the moved patient and arrow to the effective assignment.

## Orphan workflow now locked

The daily `Ορφανοί` list is an operational queue. Every unresolved/displaced patient for the date must end as either:

- placed in a valid therapist/time slot, or
- `Αναστολή θεραπείας` for that date only.

`Αναστολή θεραπείας` is a daily overlay/status, never a recurring delete, and must feed daily statistics as a non-performed treatment with that explicit reason.

No orphan may silently disappear from the plan.

## Daily operational rules locked

- capacity/load is measured by occupied timeslots,
- shared slots may host multiple patients but consume one timeslot,
- cancellations release same-day capacity,
- therapist absences and cancellations may arrive in any order,
- daily recomputation must be idempotent,
- displaced existing patients precede new-patient allocation,
- remaining orphans must be explicitly placed or suspended,
- recurring/base schedules are never destructively edited for one-day operational facts,
- date-specific moves are reversible through the daily ledger,
- mixed inpatient/outpatient presentation ambiguity must not block a valid movement.

## Accepted MASTER UX

Native V81 behavior remains green:

- four action buttons horizontally at the top,
- treatment/column headers visible,
- rows 1-4 frozen during vertical scrolling,
- no UserForm/sidebar/CommandBars workaround.

## Outpatient architecture already in place

The branch contains:

- inpatient/outpatient patient types,
- separate authoritative `OUTPATIENT_SCHEDULE`,
- unified base-schedule reads,
- outpatient presentation rules,
- daily operational integration,
- cancellation/no-show overlays,
- outpatient schedule registration/suggestion/conflict infrastructure.

Outpatient recurring create/edit native verification remains open after the current daily-state integration work.

## Immediate roadmap

1. Wire MASTER timetable to concrete-date `DailySessionState`.
2. Verify active move appears on MASTER as `Μηλιδάκης 10:45`.
3. Verify restore returns MASTER to `Σκαρώνης 09:15`.
4. Fix THERAPIST_DAILY time-column width/format.
5. Fix infectious vs robotic styling for moved patients.
6. Remove legacy `ΑΝΑΝΕΩΣΗ` button.
7. Implement/verify orphan final outcomes and statistics.
8. Test selected/all restore with multiple changes.
9. Replace preview snapshot restore with authoritative overlay removal + full daily recomputation.
10. Resume outpatient recurring create/edit native gate.
11. Continue remaining daily operational native checks and branch/PR hardening.

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
