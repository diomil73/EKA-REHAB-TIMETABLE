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
- consolidated orphan operational gate: **`43 passed, 3 warnings in 5.15s`**
- V82 individual ledger restore: native green
- V83 MASTER active-overlay sync + restore: native green
- V84 time-axis display + restore: native green
- V85 stale-orange cleanup + restore: native green

Python 3.14 `ZipFile.__del__` warnings remain non-blocking technical debt. The 43-test gate produced only this already-known warning class.

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

Implemented and verified:

- `SessionCancellationKind.TREATMENT_SUSPENDED`,
- `rehab_core/orphan_resolution.py`,
- exact reason `Αναστολή θεραπείας`,
- `rehab_core/daily_statistics.py`,
- statistics distinguish performed/replaced/absent/postponed/no-show/suspended/unresolved,
- day finalized only when unresolved reaches zero.

### Operational workflow: source green

The sequential daily replacement tool now enforces final orphan resolution:

- old `0 = leave unresolved` path removed,
- candidate number means placement,
- `S` means date-specific `Αναστολή θεραπείας`,
- `Q` aborts the whole preview,
- if no candidate exists, only `S` or `Q` is accepted,
- before workbook output, `calculate_daily_treatment_statistics()` must report `is_finalized=True`, otherwise the workflow safety-stops.

The consolidated Windows source gate covering reader/dropdown/action flow/persistence/statistics/state/queue/capacity is now green: **43 passed, 3 warnings in 5.15s**.

### Authoritative suspension persistence: source green, native pending

Generated suspension choices are persisted into the **output copy only** in `DAILY_INPUT` as a date-specific operational fact. `daily_input_reader.py` recognizes `Αναστολή θεραπείας` as `TREATMENT_SUSPENDED`, and the DAILY_INPUT dropdown includes it.

### Workbook statistics surface: source green, native pending

`Python/rehab_excel/daily_statistics_writer.py` writes visible `DAILY_STATISTICS` from the exact same finalized `DailySessionState`, including scheduled/performed/replaced/absent/postponed/no-show/suspended/unresolved/finalized.

### Native mixed-orphan smoke helper: implemented, source gate pending

New tool:

- `Python/tools/prepare_orphan_smoke_input.py`

It safely creates `Excel/previews/ORPHAN_SMOKE_INPUT_V86.xlsm` from V85, automatically chooses a therapist with at least two sessions on the selected date, writes one whole-day therapist absence into DAILY_INPUT, keeps the source unchanged, and reports the affected sessions. This gives a deterministic native case where one orphan can be placed and another can be suspended.

Regression file:

- `Tests/test_prepare_orphan_smoke_input.py`

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

1. Run the small source gate for the new native smoke preparer.
2. Build `ORPHAN_SMOKE_INPUT_V86.xlsm` for `2026-10-19`.
3. Run `apply_daily_input_replacements.py` to `ORPHAN_RESOLUTION_V87.xlsm` and choose at least one placement plus at least one `S = Αναστολή θεραπείας`.
4. Verify native output: placement visible, suspended patient removed from effective schedule, output DAILY_INPUT contains suspension fact, DAILY_STATISTICS has suspended count + unresolved=0 + finalized=ΝΑΙ, source unchanged and VBA preserved.
5. Re-read the output to confirm suspension persistence.
6. Then test selected/all restore with multiple changes.
7. Replace snapshot restore with authoritative overlay removal + recompute.
8. Resume outpatient recurring create/edit native gate.
9. Clean Python 3.14 ZipFile teardown warnings later.

Next Windows commands:

```powershell
python -m pytest Tests/test_prepare_orphan_smoke_input.py -q

python Python\tools\prepare_orphan_smoke_input.py `
  --source ".\Excel\previews\DAILY_CHANGE_LEDGER_V85.xlsm" `
  --output ".\Excel\previews\ORPHAN_SMOKE_INPUT_V86.xlsm" `
  --date 2026-10-19 `
  --overwrite
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
