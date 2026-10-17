# Project Status

Last updated: 2026-10-17

## Current stage

The project remains on branch `feature/outpatient-scheduling`.

The patient-centric multi-treatment scheduling milestone is closed. The active operational milestone is **one concrete-date effective state across MASTER + THERAPIST_DAILY + ledger + statistics**, with outpatient recurring create/edit still open behind it.

Latest full-suite checkpoint after cleanup: **699 passed, 19 warnings**.
Latest focused operational regression checkpoint: **31 passed in 2.04s**.

## Current native checkpoint

`Excel/previews/DAILY_CHANGE_LEDGER_V82.xlsm` builds successfully and individual restore is native green.

Stress case:

- date `2026-10-19`,
- patient `ΜΕΝΤΗΣ`,
- original `Σκαρώνης 09:15`,
- selected replacement `Μηλιδάκης 10:45`,
- candidate load `load 5/6 | patients 6`,
- source unchanged,
- VBA preserved,
- ledger installed,
- per-row restore returns the patient correctly to the original THERAPIST_DAILY position.

## MASTER daily-state integration

Confirmed native bug: while the move was active, `THERAPIST_DAILY` showed the replacement but MASTER remained on the recurring assignment.

Source fix is now implemented, pending focused tests and native verification:

- `Python/rehab_excel/master_daily_overlay.py` derives MASTER treatment-cell patches from `DailySessionState`,
- guided replacement composes that overlay into the same preview plan,
- recurring source stays unchanged,
- ledger restore invokes the existing MASTER patient-program refresh so the preview can return MASTER to recurring state,
- dedicated regression coverage was added.

Required invariant:

`DailySessionState -> THERAPIST_DAILY + MASTER timetable + ledger + statistics`

Acceptance still pending:

- active MASTER display `Μηλιδάκης 10:45`,
- restore MASTER display `Σκαρώνης 09:15`.

## V82 native cleanup backlog

- left time column in `THERAPIST_DAILY` shows `###`; widen/fix `hh:mm`,
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

1. Run focused MASTER daily-overlay + ledger restore-hook tests.
2. Build fresh guided replacement preview.
3. Verify active MASTER shows `Μηλιδάκης 10:45`.
4. Verify restore returns MASTER and THERAPIST_DAILY to `Σκαρώνης 09:15`.
5. Fix THERAPIST_DAILY time-column width/format.
6. Fix infectious vs robotic styling for moved patients.
7. Remove legacy `ΑΝΑΝΕΩΣΗ` button.
8. Implement/verify orphan final outcomes and statistics.
9. Test selected/all restore with multiple changes.
10. Replace preview snapshot restore with authoritative overlay removal + full daily recomputation.
11. Resume outpatient recurring create/edit native gate.
12. Continue remaining native checks and branch/PR hardening.

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
