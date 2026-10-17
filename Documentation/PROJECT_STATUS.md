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
- consolidated orphan operational gate: **43 passed, 3 warnings in 5.15s**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## Current native checkpoint

`Excel/previews/DAILY_CHANGE_LEDGER_V85.xlsm` remains the latest fully accepted native checkpoint.

Confirmed native green for the ΜΕΝΤΗΣ stress case:

- active replacement appears consistently in THERAPIST_DAILY and MASTER,
- source trace remains visible,
- individual restore returns both views to original,
- time axis displays `hh:mm` with no `####`,
- moved non-robotic ΜΕΝΤΗΣ no longer inherits robotic-orange styling.

## Deferred cosmetic issue

The obsolete `ΑΝΑΝΕΩΣΗ` Form Control remains visible. Decision: **defer/manual cleanup**. It is not a functional blocker.

## Orphan workflow status

Every patient in `Ορφανοί` for the concrete date must resolve as either placement or `Αναστολή θεραπείας` for that date only.

Operational implementation is source green:

- mandatory placement/suspension flow,
- no unresolved completion path,
- finalization gate requires unresolved=0,
- suspension persists into output DAILY_INPUT,
- reread recognizes `TREATMENT_SUSPENDED`,
- DAILY_INPUT dropdown includes `Αναστολή θεραπείας`,
- DAILY_STATISTICS writes from the same finalized daily state,
- consolidated verification: **43 passed, 3 warnings in 5.15s**.

## Native V87 blocker and fix

The first mixed-orphan V87 run stopped safely before workbook output because `outpatient_presentation` emitted a stale-fill cleanup patch for `THERAPIST_DAILY!A21`, outside the confirmed write zones `B2:J8` and `B12:J18`.

This was a stale-colour discovery bug, not a missing layout zone. The matching green fill at A21 was unrelated sheet furniture.

Fix implemented:

- outpatient stale-fill discovery and cleanup now operate only inside the two confirmed THERAPIST_DAILY patient grids,
- write zones remain unchanged and narrow,
- regression explicitly checks that `A21` is ignored.

Commits:

- `5c81ce40256ddd36379d838c35b2a62debf0ef95`
- `7e4031fd9f01b4b11233b6bc087078b4f94686aa`

Status: **focused source gate pending, then rerun V87 from existing V86**.

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

Native green: active move rendering, MASTER sync, source trace, per-row restore, time-axis display, stale-orange cleanup.

Still open: selected/all restore with multiple changes and replacing snapshot restore with authoritative overlay removal + full daily recomputation.

## Outpatient architecture

The branch already contains inpatient/outpatient types, separate authoritative `OUTPATIENT_SCHEDULE`, unified base reads, outpatient presentation, cancellation/no-show overlays and recurring schedule infrastructure.

Outpatient recurring create/edit native verification remains open after the current orphan/daily-state milestone.

## Immediate roadmap

1. Run focused tests for outpatient presentation + orphan operational flow after A21 fix.
2. Re-run `apply_daily_input_replacements.py` from `ORPHAN_SMOKE_INPUT_V86.xlsm` to `ORPHAN_RESOLUTION_V87.xlsm`.
3. Include at least one placement and one suspension.
4. Verify output DAILY_INPUT persistence, DAILY_STATISTICS unresolved=0/finalized=ΝΑΙ, source unchanged and VBA preserved.
5. Re-read V87 to confirm suspension persistence.
6. Test selected/all restore with multiple changes.
7. Replace snapshot restore with authoritative overlay removal + full recompute.
8. Resume outpatient recurring create/edit native gate.
9. Clean Python 3.14 ZipFile teardown warnings later.

## Safety constraints

- Never write directly to the baseline `.xlsm` during development or smoke testing.
- Preview workflows operate on copied `.xlsm` files only.
- Preserve/verify VBA when rebuilding previews.
- Keep layout write zones narrow; fix producer logic rather than expanding zones without evidence.
- Keep current scheduling behavior backward-compatible unless an explicit business rule changes it.
- Windows test command is always `python -m pytest`.

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests
