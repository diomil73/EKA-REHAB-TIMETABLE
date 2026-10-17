# Current Handoff

Last updated: 2026-10-17

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Latest native preview tested: `Excel/previews/DAILY_CHANGE_LEDGER_V85.xlsm`
- V81 MASTER toolbar/header UX remains native green.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- replacement load + mixed outpatient/inpatient regression: `31 passed in 2.04s`
- MASTER daily overlay + ledger restore source gate: `24 passed, 5 warnings in 4.39s`
- V82 individual ledger restore: native green
- V83 MASTER active-overlay sync + restore: native green
- V84 time-axis display + restore: native green
- V85 stale-orange cleanup + restore: native green

Python 3.14 `ZipFile.__del__` warnings remain non-blocking technical debt.

## Native stress case now green

Case:

- date `2026-10-19`
- patient `ΜΕΝΤΗΣ`
- original `Σκαρώνης 09:15`
- replacement `Μηλιδάκης 10:45`

Confirmed native:

- `THERAPIST_DAILY` and `MASTER` agree on the effective daily assignment,
- source-position trace remains visible and is accepted UX,
- individual `Επαναφορά` returns both views to the original assignment,
- left-side time axis shows readable `hh:mm` with no `####`,
- moved non-robotic ΜΕΝΤΗΣ no longer inherits stale robotic-orange styling.

## Legacy `ΑΝΑΝΕΩΣΗ` button: deferred/manual cleanup

The obsolete Form Control still remains despite targeted source cleanup attempts.

Native diagnostic identified it as:

- Form Control / Shape type `8`,
- name `Button 1`,
- caption `ΑΝΑΝΕΩΣΗ`,
- macro `...!RefreshAll`.

Decision: **do not spend more engineering time on this now**. It is cosmetic, not a functional blocker. The user will remove it manually in Excel later if needed. Do not reopen this item unless it becomes operationally relevant.

## ACTIVE development focus: orphan workflow + statistics

The domain foundation for the daily `Ορφανοί` workflow is now implemented in source and awaits the focused source gate.

Locked rule: every orphan for the concrete date must resolve as exactly one of:

- placed into a valid therapist/time slot, or
- `Αναστολή θεραπείας` for that date only.

Architecture decision:

- do **not** add a parallel `DailySessionStatus` for suspension,
- represent suspension as `DailySessionStatus.CANCELLED` plus `SessionCancellationKind.TREATMENT_SUSPENDED`,
- this preserves the existing cancellation precedence, releases the provider slot, keeps recurring/base untouched, and carries an auditable reason.

Implemented domain slice:

- `SessionCancellationKind.TREATMENT_SUSPENDED`,
- `Python/rehab_core/orphan_resolution.py`,
- `create_treatment_suspension(session)` with exact reason `Αναστολή θεραπείας`,
- `Python/rehab_core/daily_statistics.py`,
- statistics distinguish performed, replaced, patient absent, department postponed, no-show, treatment suspended and unresolved,
- remaining `THERAPIST_ABSENT` sessions count as unresolved rather than silently cancelled,
- the day is finalized only when unresolved sessions reach zero,
- focused tests added in `Tests/test_orphan_treatment_suspension.py` and `Tests/test_daily_treatment_statistics.py`.

Status: **implemented, source gate pending; Excel orphan actions/UI not yet wired**.

## Locked daily recomputation order

1. recurring/base schedule for concrete date,
2. cancellations/no-shows release capacity,
3. therapist absences,
4. identify displaced/unscheduled existing patients,
5. redistribute into net free capacity,
6. resolve remaining orphans by placement or `Αναστολή θεραπείας`,
7. only then expose remaining capacity to new patients.

Recomputation must remain idempotent and order-independent.

## Reversible daily change ledger

Native green:

- active move rendering,
- MASTER synchronization,
- source trace,
- per-row restore,
- time-axis display,
- stale-orange cleanup.

Still open:

- selected/all restore with multiple simultaneous changes,
- production restore via authoritative overlay removal + full daily recompute instead of snapshot restoration.

## Exact resume point

Priority order:

1. **Run the focused source gate for treatment suspension + statistics + existing daily-state/queue/capacity regressions.**
2. If green, wire orphan operational actions into the actual workflow: place or suspend.
3. Persist `Αναστολή θεραπείας` as an authoritative date-specific daily fact, not preview-only state.
4. Surface daily statistics in the workbook/reporting flow.
5. Native-check orphan placement and suspension behavior.
6. Test selected/all restore with multiple changes.
7. Replace snapshot restore with authoritative overlay removal + recompute.
8. Resume outpatient recurring create/edit native gate.
9. Clean Python 3.14 ZipFile teardown warnings later.

Focused Windows gate:

```powershell
python -m pytest `
  Tests/test_orphan_treatment_suspension.py `
  Tests/test_daily_treatment_statistics.py `
  Tests/test_daily_state.py `
  Tests/test_therapist_absence_queue.py `
  Tests/test_cancellation_releases_replacement_capacity.py -q
```

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
