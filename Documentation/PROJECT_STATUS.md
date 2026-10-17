# Project Status

Last updated: 2026-10-17

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **orphan resolution + date-specific treatment suspension + correct daily statistics**, on top of the shared concrete-date effective state used by MASTER, THERAPIST_DAILY and the ledger.

Latest confirmed checkpoints:

- full regression: **699 passed, 19 warnings**
- replacement load + mixed-slot gate: **31 passed in 2.04s**
- MASTER daily-overlay + ledger restore gate: **24 passed, 5 warnings in 4.39s**
- orphan suspension + statistics domain gate: **21 passed in 3.11s**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## Current native checkpoint

`Excel/previews/DAILY_CHANGE_LEDGER_V85.xlsm` remains the latest native checkpoint.

Confirmed native green for the ΜΕΝΤΗΣ stress case:

- active replacement appears consistently in THERAPIST_DAILY and MASTER,
- source trace remains visible,
- individual restore returns both views to original,
- time axis displays `hh:mm` with no `####`,
- moved non-robotic ΜΕΝΤΗΣ no longer inherits robotic-orange styling.

## Deferred cosmetic issue

The obsolete `ΑΝΑΝΕΩΣΗ` Form Control remains visible. Decision: **defer/manual cleanup**. It is not a functional blocker.

## Orphan workflow status

Every patient in `Ορφανοί` for the concrete date must resolve as either:

- placed in a valid therapist/time slot, or
- `Αναστολή θεραπείας` for that date only.

Architecture:

- `Αναστολή θεραπείας` = `DailySessionStatus.CANCELLED` + `SessionCancellationKind.TREATMENT_SUSPENDED`,
- recurring/base remains untouched,
- provider slot is released,
- explicit reason remains auditable,
- finalized day requires unresolved=0.

Domain foundation is **source green: 21 passed in 3.11s**.

New operational implementation, awaiting consolidated source/native gate:

- `apply_daily_input_replacements.py` now requires placement or suspension for every orphan; old `leave unresolved` path removed,
- `S` records suspension, `Q` aborts the preview, and no-candidate cases require S/Q,
- finalization safety gate blocks output when unresolved sessions remain,
- `daily_input_reader.py` recognizes `Αναστολή θεραπείας`,
- DAILY_INPUT status dropdown includes `Αναστολή θεραπείας`,
- `daily_operational_fact_writer.py` persists generated suspensions into output DAILY_INPUT so the output can reproduce the same fact on reread,
- `daily_statistics_writer.py` writes a visible `DAILY_STATISTICS` sheet from the finalized daily state,
- terminal summary and workbook statistics come from the same statistics object.

## Daily operational rules locked

1. recurring/base schedule for the concrete date,
2. cancellations/no-shows release same-day capacity,
3. therapist absences,
4. identify displaced/unscheduled existing patients,
5. redistribute existing patients into net free capacity,
6. resolve remaining orphans by placement or `Αναστολή θεραπείας`,
7. only then expose remaining capacity to new patients.

Operational capacity remains timeslot-based. Recurring/base schedules are never destructively edited by one-day facts. Recompute must remain idempotent and order-independent.

## Ledger status

Native green:

- active move rendering,
- MASTER sync,
- source trace,
- per-row restore,
- time-axis display,
- stale-orange cleanup.

Still open:

- selected/all restore with multiple changes,
- replacing preview snapshot restore with authoritative overlay removal + full daily recomputation.

## Outpatient architecture

The branch already contains inpatient/outpatient types, separate authoritative `OUTPATIENT_SCHEDULE`, unified base reads, outpatient presentation, cancellation/no-show overlays and recurring schedule infrastructure.

Outpatient recurring create/edit native verification remains open after the current orphan/daily-state milestone.

## Immediate roadmap

1. Run consolidated source gate for orphan action flow + suspension persistence + DAILY_INPUT dropdown + DAILY_STATISTICS.
2. Run native mixed orphan scenario with at least one placement and one treatment suspension.
3. Verify output reread, stats finalization, source unchanged and VBA preserved.
4. Test selected/all restore with multiple changes.
5. Replace snapshot restore with authoritative overlay removal + full recompute.
6. Resume outpatient recurring create/edit native gate.
7. Continue branch/PR hardening.
8. Clean Python 3.14 ZipFile teardown warnings later.

Focused Windows gate:

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
