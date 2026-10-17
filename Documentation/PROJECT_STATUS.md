# Project Status

Last updated: 2026-10-17

## Current stage

The project remains on branch `feature/outpatient-scheduling`.

The active operational milestone is **one concrete-date effective state across MASTER + THERAPIST_DAILY + ledger + statistics**, with orphan resolution and outpatient recurring create/edit still open behind it.

Latest full-suite checkpoint after cleanup: **699 passed, 19 warnings**.
Latest focused operational regression checkpoint: **31 passed in 2.04s**.
Latest MASTER daily-overlay gate: **24 passed, 5 warnings in 4.39s**.

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## Current native checkpoint

`Excel/previews/DAILY_CHANGE_LEDGER_V85.xlsm` is the latest native checkpoint.

Stress case:

- date `2026-10-19`,
- patient `ΜΕΝΤΗΣ`,
- original `Σκαρώνης 09:15`,
- replacement `Μηλιδάκης 10:45`.

Native green:

- THERAPIST_DAILY effective move + source trace,
- MASTER effective daily assignment,
- individual restore returns both views to original,
- readable left time axis with no `####`,
- moved ΜΕΝΤΗΣ no longer inherits robotic-orange styling.

## Remaining presentation issue: legacy refresh button

The obsolete `ΑΝΑΝΕΩΣΗ` button still appeared in V85. Native diagnostic identified it exactly as:

- Excel Form Control / Shape type `8`,
- name `Button 1`,
- caption `ΑΝΑΝΕΩΣΗ`,
- macro `DAILY_CHANGE_LEDGER_V85.xlsm!RefreshAll`,
- OLEObjects `0`.

A surgical cleanup is now implemented:

- enumerate `ws.Buttons()` directly,
- remove only when name=`Button 1`, caption=`ΑΝΑΝΕΩΣΗ`, and terminal macro=`RefreshAll`,
- keep a shape fallback with the same triple identity,
- ledger and toolbar controls are explicitly outside the target pattern.

Commits:

- `e62ce5354752a61c4b10568ba184991b51b24c67` regression
- `5c51e8988797ce67358a6886e5e41f9bf5bb848c` implementation

Status: **source implemented, focused/native verification pending**.

## Orphan workflow locked

Every patient in `Ορφανοί` for the concrete date must end as either:

- placed in a valid therapist/time slot, or
- `Αναστολή θεραπείας` for that date only.

Suspension is a daily overlay/status, never a recurring delete, and must feed daily statistics as a non-performed treatment with explicit reason. Existing displaced patients precede new-patient allocation.

## Daily operational rules locked

- capacity/load is occupied timeslots,
- shared slots may host multiple patients but consume one timeslot,
- cancellations release same-day capacity,
- therapist absences and cancellations may arrive in any order,
- recomputation is idempotent,
- remaining orphans are explicitly placed or suspended,
- recurring/base schedules are never destructively edited for one-day facts,
- date-specific moves are reversible,
- MASTER and THERAPIST_DAILY must agree on effective daily state.

## Accepted MASTER UX

Native V81 remains green:

- four action buttons horizontally at the top,
- treatment/column headers visible,
- rows 1-4 frozen during vertical scrolling,
- no UserForm/sidebar/CommandBars workaround.

## Outpatient architecture already in place

The branch contains inpatient/outpatient types, separate authoritative `OUTPATIENT_SCHEDULE`, unified base reads, outpatient presentation, daily operational integration, cancellation/no-show overlays, and outpatient schedule registration/suggestion/conflict infrastructure.

Outpatient recurring create/edit native verification remains open after the current daily-state/orphan work.

## Immediate roadmap

1. Run focused legacy-button regression.
2. Build fresh guided replacement preview.
3. Verify legacy `ΑΝΑΝΕΩΣΗ` is gone while toolbar/ledger controls remain.
4. Close presentation cleanup milestone if green.
5. Implement orphan placement / `Αναστολή θεραπείας` outcomes and statistics.
6. Test selected/all restore with multiple changes.
7. Replace preview snapshot restore with authoritative overlay removal + full daily recomputation.
8. Resume outpatient recurring create/edit native gate.
9. Continue remaining native checks and branch/PR hardening.
10. Clean Python 3.14 ZipFile teardown warnings when convenient.

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
