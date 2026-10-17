# Current Handoff

Last updated: 2026-10-17

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested: `Excel/previews/DAILY_CHANGE_LEDGER_V85.xlsm`
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
- MASTER daily overlay + ledger restore source gate: **24 passed, 5 warnings in 4.39s**
- V82 individual ledger restore: **native green**
- V83 MASTER active-overlay sync + restore: **native green**
- V84 time-axis display + restore: **native green**
- V85 stale-orange cleanup + restore: **native green**

The 5 warnings are Python 3.14 `ZipFile.__del__` teardown warnings. They remain non-blocking technical debt.

## V85 native checkpoint

Stress case:

- date: `2026-10-19`
- patient: `ΜΕΝΤΗΣ`
- original: `Σκαρώνης 09:15`
- replacement: `Μηλιδάκης 10:45`

Native behavior now confirmed:

- `THERAPIST_DAILY` and `MASTER` agree on the effective daily assignment,
- left-side times display correctly as `hh:mm`,
- ΜΕΝΤΗΣ no longer inherits stale robotic-orange text at the Μηλιδάκης destination,
- individual `Επαναφορά` still works correctly and returns both views to `Σκαρώνης 09:15`,
- source-position trace remains accepted UX.

## Exact legacy `ΑΝΑΝΕΩΣΗ` identity

V85 still showed the obsolete refresh button. Native control inventory removed the ambiguity:

- type: Excel Form Control button / Shape type `8`,
- shape/button name: `Button 1`,
- caption/text/alt: `ΑΝΑΝΕΩΣΗ`,
- macro: `...!RefreshAll`,
- OLEObjects: `0`.

Therefore this is not ActiveX. The earlier generic shape-caption cleanup missed it because the COM caption path for this Form Control is inconsistent.

Surgical source fix now added:

- enumerate `ws.Buttons()` directly,
- delete only when all three identifiers match: `Button 1` + `ΑΝΑΝΕΩΣΗ` + macro `RefreshAll`,
- keep a shape fallback using the same triple identity,
- do not target any `ekaLedger_*` or toolbar control.

Commits:

- `e62ce5354752a61c4b10568ba184991b51b24c67` regression for exact Form Control identity
- `5c51e8988797ce67358a6886e5e41f9bf5bb848c` surgical `ws.Buttons()` cleanup

Status: **implemented, focused/native verification pending**.

## Locked orphan workflow

Every patient in the daily `Ορφανοί` queue must resolve for the concrete date as either:

- placed into a valid therapist/time slot, or
- `Αναστολή θεραπείας` for that date only.

Rules:

- never silently disappear,
- suspension is a daily overlay/status, not a recurring delete,
- suspension feeds daily statistics as treatment not performed with explicit reason,
- displaced existing patients have priority over new-patient allocation.

## Locked daily recomputation order

1. recurring/base schedule for concrete date,
2. cancellations/no-shows release capacity,
3. therapist absences,
4. identify displaced/unscheduled existing patients,
5. redistribute into net free capacity,
6. resolve remaining orphans by placement or `Αναστολή θεραπείας`,
7. only then expose remaining capacity to new patients.

Recomputation must be idempotent and recurring schedules are never destructively edited by one-day facts.

## Reversible daily change ledger

Native green so far:

- active move rendering,
- MASTER synchronization,
- source trace,
- per-row restore,
- time-axis display,
- stale-orange cleanup.

Still open:

- selected/all restore with multiple changes,
- production restore as authoritative overlay removal + full daily recompute instead of snapshot restoration.

## Current development focus / exact resume point

Priority order:

1. **Run focused regression for exact legacy Form Control removal.**
2. Build fresh preview and verify `ΑΝΑΝΕΩΣΗ` is gone while ledger/toolbar controls remain.
3. If green, close presentation cleanup milestone.
4. Implement orphan queue semantics: placement or date-specific `Αναστολή θεραπείας`.
5. Integrate suspension into daily statistics.
6. Test selected/all restore with multiple changes.
7. Replace snapshot restore with authoritative overlay removal + recompute.
8. Resume outpatient recurring create/edit native gate.
9. Clean Python 3.14 ZipFile teardown warnings later.

## Business rules that must not regress

- ΦΘ provider is required and solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ provider/name is optional free text and not a solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ have no provider/name field.
- Outpatients cannot receive Ρομποτικό.
- Nonblank manual constraints are never silently overwritten.
- Standard duration 45 minutes; Ανακλινόμενο 30 minutes.
- Exact adjacency is legal.
- `PATIENT_PLANNER` authoritative for inpatient recurring; `OUTPATIENT_SCHEDULE` authoritative for outpatient recurring.
- one-day cancellations, absences, replacements, suspensions and redistributions are overlays, not recurring edits,
- operational load/capacity is timeslot-based,
- existing displaced patients precede new patients,
- unresolved orphans must be placed or explicitly suspended,
- `MASTER` and `THERAPIST_DAILY` must agree on effective daily state.

## Recovery order if context is lost

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.
