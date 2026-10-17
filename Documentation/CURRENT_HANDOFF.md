# Current Handoff

Last updated: 2026-10-17

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested: `MASTER_TOOLBAR_SMOKE_V81.xlsm`
- Latest successfully generated ledger preview: `Excel/previews/DAILY_CHANGE_LEDGER_V82.xlsm`
- V73 scheduling/conflict behavior remains native green.
- V80 clean reset preview is native green.
- V81 MASTER toolbar/header pinning is native green.

## Latest confirmed test checkpoints

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`
- cleanup/reset focused gate: `42 passed in 3.70s`
- row-4 toolbar/header pinning focused source gate: `43 passed in 1.92s`
- full regression after cleanup: `699 passed, 19 warnings`
- reversible daily change ledger focused source gate: `10 passed in 3.26s`
- replacement load + mixed outpatient/inpatient regression gate: `31 passed in 2.04s`
- V82 guided replacement native build: green
- V82 individual ledger restore: **native green**

## V82 native guided replacement checkpoint

Real stress case:

- date: `2026-10-19`
- patient: `ΜΕΝΤΗΣ`
- original assignment: `Σκαρώνης 09:15`
- selected replacement: `Μηλιδάκης 10:45`
- candidate display for Μηλιδάκης: `load 5/6 | patients 6`
- source workbook unchanged: `True`
- VBA preserved: `True`
- daily change ledger installed: `yes`
- output: `Excel/previews/DAILY_CHANGE_LEDGER_V82.xlsm`

Native UX findings:

- source cell intentionally keeps ΜΕΝΤΗΣ visible with a movement arrow/note; user likes this and it should be retained,
- destination shows the moved patient,
- per-row `Επαναφορά` works correctly and returns the patient to the original position,
- individual restore is therefore native green,
- selected/all restore remains to be tested once multiple simultaneous preview changes are supported.

## NEW highest-priority integration gap: MASTER timetable does not follow daily effective state

Native verification showed that while the replacement was active, `THERAPIST_DAILY` showed ΜΕΝΤΗΣ at `Μηλιδάκης 10:45`, but the timetable on `MASTER` did **not** change from the original recurring assignment.

This is a confirmed bug/integration gap, not accepted behavior.

Locked rule:

- `MASTER` timetable and `THERAPIST_DAILY` must project the same concrete-date `DailySessionState`,
- while a daily move is active, MASTER must show the effective therapist/time,
- after restore, MASTER must return to the original therapist/time,
- MASTER must not keep a separate recurring-only truth when the selected/operational date has daily overlays.

Target architecture:

`DailySessionState -> THERAPIST_DAILY + MASTER timetable + ledger + statistics`

This synchronization is now the first implementation task before cosmetic cleanup.

## Native cleanup items discovered in V82

The user identified these concrete UI issues in `THERAPIST_DAILY`:

1. left time column is too narrow / some times render as `###`; widen and lock `hh:mm` display,
2. ΜΕΝΤΗΣ received orange text incorrectly; orange is reserved for robotic participation only,
3. ΜΕΝΤΗΣ is infectious and not robotic, so infectious presentation must win and no robotic-orange text may leak from destination/base styling,
4. obsolete right-side VBA-era button `ΑΝΑΝΕΩΣΗ` remains and must be removed.

Keep the useful source-position visual trace for moved patients.

## Locked orphan workflow

Do not forget the daily list of unresolved/displaced patients (`Ορφανοί`). Every orphan for the concrete date must end in one of two operational outcomes:

- **placed** into a valid therapist/time slot, or
- **removed for that date only** with reason `Αναστολή θεραπείας`.

Rules:

- this is a daily overlay/status, never a destructive edit of the recurring/base schedule,
- an orphan must not silently disappear from the operational plan,
- `Αναστολή θεραπείας` must feed daily statistics as a treatment not performed for that explicit reason,
- orphan resolution participates in the same recomputable daily-state engine,
- displaced existing patients have priority over new-patient allocation.

## Locked daily-operations rule: cancellations + therapist absences

Patient cancellations for both inpatients and outpatients release same-day timeslots. Daily computation must be independent of entry order:

1. recurring/base schedule for selected date,
2. apply cancellations/no-shows and release slots,
3. apply therapist absences,
4. identify displaced/unscheduled patients,
5. redistribute existing affected patients into net free capacity,
6. resolve any remaining orphans by placement or `Αναστολή θεραπείας`,
7. only then expose remaining capacity to new patients.

Recomputation must be idempotent and recurring schedules must never be destructively edited for one-day facts.

## Locked replacement load semantics

Operational therapist load is occupied daily **timeslots**, not patient/session count.

- `active_timeslots` is authoritative load/capacity,
- multiple patients sharing one clock slot count as multiple sessions but one occupied timeslot,
- replacement ranking sorts primarily by occupied timeslots,
- candidate display distinguishes `load x/y` from `patients n`.

## Mixed inpatient/outpatient presentation rule

A mixed inpatient/outpatient effective slot must not block a valid operational preview. Do not apply one whole-cell outpatient fill to a mixed cell; keep it neutral until per-patient styling exists.

## Reversible daily change ledger

Required section in `THERAPIST_DAILY`:

`Αλλαγές για την dd/mm/yy`

Each move preserves patient, original/effective therapist and time, reason/source, stable id and status.

Controls:

- per-row `Επαναφορά`,
- selection,
- `Επαναφορά επιλεγμένων`,
- `Επαναφορά όλων`.

Production restore semantics:

- remove/reverse the date-specific override and recompute the day,
- never reconstruct or rewrite recurring schedule from UI text,
- idempotent,
- stale undo blocked,
- recalculate capacity/conflicts/replacements after restore,
- bulk restore validated and safe.

Prototype currently uses preview snapshots only; after UX acceptance it must be connected to authoritative overlay removal + full daily recomputation.

## Current development focus / exact resume point

Priority order:

1. **Wire MASTER timetable to the same daily effective-state projection as THERAPIST_DAILY.**
2. Native test active move: ΜΕΝΤΗΣ must show `Μηλιδάκης 10:45` in MASTER while the overlay is active.
3. Native test restore: MASTER must return to `Σκαρώνης 09:15` after `Επαναφορά`.
4. Fix THERAPIST_DAILY time column width/format.
5. Fix robotic-orange vs infectious styling precedence for moved patients.
6. Remove legacy `ΑΝΑΝΕΩΣΗ` button.
7. Implement/verify orphan queue semantics, including `Αναστολή θεραπείας` and statistics.
8. Test selected/all restore with multiple changes.
9. Replace preview snapshot restore with authoritative overlay removal + daily recompute.
10. Resume outpatient recurring create/edit native gate.

## Business rules that must not regress

- ΦΘ provider is required and solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ provider/name is optional free text and not a solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ have no provider/name field.
- Outpatients cannot receive Ρομποτικό.
- Nonblank manual constraints are never silently overwritten.
- Standard duration 45 minutes; Ανακλινόμενο 30 minutes.
- Exact adjacency is legal.
- `PATIENT_PLANNER` authoritative for inpatient recurring schedules; `OUTPATIENT_SCHEDULE` authoritative for outpatient recurring schedules.
- one-day cancellations, absences, replacements, suspensions and redistributions are daily overlays, not destructive recurring edits,
- operational load/capacity is timeslot-based,
- existing displaced patients precede new patients,
- unresolved orphans must be placed or explicitly suspended for the date,
- `MASTER` and `THERAPIST_DAILY` must agree on the effective daily state.

## Recovery order if context is lost

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.
