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
- orphan suspension + statistics domain gate: `21 passed in 3.11s`
- V82 individual ledger restore: native green
- V83 MASTER active-overlay sync + restore: native green
- V84 time-axis display + restore: native green
- V85 stale-orange cleanup + restore: native green

Python 3.14 `ZipFile.__del__` warnings remain non-blocking technical debt.

## Native stress case already green

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

The obsolete Form Control remains visible. Decision: **do not spend more engineering time on this now**. It is cosmetic, not a functional blocker. The user will remove it manually in Excel later if needed.

## ACTIVE development focus: orphan workflow + statistics

Locked rule: every orphan for the concrete date must resolve as exactly one of:

- placed into a valid therapist/time slot, or
- `Αναστολή θεραπείας` for that date only.

Architecture:

- no parallel daily status for suspension,
- suspension is `DailySessionStatus.CANCELLED` plus `SessionCancellationKind.TREATMENT_SUSPENDED`,
- recurring/base stays untouched,
- original provider slot is released,
- reason is auditable.

### Domain foundation: source green

Implemented and previously verified with **21 passed in 3.11s**:

- `SessionCancellationKind.TREATMENT_SUSPENDED`,
- `rehab_core/orphan_resolution.py`,
- exact reason `Αναστολή θεραπείας`,
- `rehab_core/daily_statistics.py`,
- statistics distinguish performed/replaced/absent/postponed/no-show/suspended/unresolved,
- day finalized only when unresolved reaches zero.

### Operational workflow: implemented, new source gate pending

The sequential daily replacement tool has now been converted into mandatory orphan resolution:

- removed the old `0 = leave unresolved` path,
- candidate number means placement,
- `S` means date-specific `Αναστολή θεραπείας`,
- `Q` aborts the whole preview,
- if no candidate exists, only `S` or `Q` is accepted,
- before workbook output, `calculate_daily_treatment_statistics()` must report `is_finalized=True`, otherwise the workflow safety-stops.

Files:

- `Python/tools/apply_daily_input_replacements.py`
- `Tests/test_apply_daily_input_replacements_resolution.py`

### Authoritative suspension persistence: implemented, native gate pending

Generated suspension choices no longer live only in Python memory.

`Python/rehab_excel/daily_operational_fact_writer.py` appends each generated suspension into the **output copy only** in `DAILY_INPUT` as:

- patient,
- `ΟΧΙ` for whole-day,
- original `HH:MM`,
- status `Αναστολή θεραπείας`,
- comment `Ορφανός - ημερήσια επιχειρησιακή απόφαση`.

The writer is idempotent for patient/time/status. Re-reading the output workbook therefore recreates the same `TREATMENT_SUSPENDED` fact. Recurring/base source is not edited.

`daily_input_reader.py` now recognizes `Αναστολή θεραπείας` as a session-specific cancellation and the generated DAILY_INPUT dropdown includes it as a selectable status.

Relevant tests:

- `Tests/test_daily_input_cancellations.py`
- `Tests/test_daily_operational_fact_writer.py`
- `Tests/test_daily_input_sheet.py`

### Workbook statistics surface: implemented, native gate pending

`Python/rehab_excel/daily_statistics_writer.py` writes a visible `DAILY_STATISTICS` sheet into the output preview from the exact same finalized `DailySessionState` used by scheduling.

It surfaces:

- scheduled,
- performed,
- replaced,
- patient absent,
- department postponed,
- patient no-show,
- treatment suspended,
- unresolved,
- finalized `ΝΑΙ/ΟΧΙ`.

`apply_daily_input_replacements.py` now writes this sheet after suspension-fact persistence.

Test: `Tests/test_daily_statistics_writer.py`.

## Locked daily recomputation order

1. recurring/base schedule for concrete date,
2. cancellations/no-shows release capacity,
3. therapist absences,
4. identify displaced/unscheduled existing patients,
5. redistribute existing affected patients into net free capacity,
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

1. **Run the consolidated source gate for the new orphan operational workflow, persistence, DAILY_INPUT dropdown and workbook statistics surface.**
2. If green, run one native mixed orphan scenario with at least one placement and one `S = Αναστολή θεραπείας`.
3. Verify output `DAILY_INPUT` can be re-read with suspension fact, `DAILY_STATISTICS` has suspended count + unresolved=0 + finalized=ΝΑΙ, source unchanged and VBA preserved.
4. Then test selected/all restore with multiple changes.
5. Replace snapshot restore with authoritative overlay removal + recompute.
6. Resume outpatient recurring create/edit native gate.
7. Clean Python 3.14 ZipFile teardown warnings later.

Next focused Windows gate:

```powershell
python -m pytest `
  Tests/test_daily_input_cancellations.py `
  Tests/test_daily_input_sheet.py `
  Tests/test_apply_daily_input_replacements_resolution.py `
  Tests/test_daily_operational_fact_writer.py `
  Tests/test_daily_statistics_writer.py `
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
