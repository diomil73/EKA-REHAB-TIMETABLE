# Project Status

Last updated: 2026-10-17

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **authoritative daily operational facts + restart-safe recomputation**.

Latest confirmed checkpoints:

- full regression: **699 passed, 19 warnings**
- replacement load + mixed-slot gate: **31 passed in 2.04s**
- MASTER daily-overlay + ledger restore gate: **24 passed, 5 warnings in 4.39s**
- orphan suspension + statistics domain gate: **21 passed in 3.11s**
- consolidated orphan operational gate: **43 passed, 3 warnings in 5.15s**
- authoritative replacement fact focused source gate: **28 passed, 3 warnings in 5.14s**
- V87 mixed orphan scenario: **native green**
- V87 suspension/statistics reread persistence: **native green**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## Native V87 result

`Excel/previews/ORPHAN_RESOLUTION_V87.xlsm`

Confirmed:

- 5 accepted replacement placements,
- 2 treatment suspensions,
- 2 persisted suspension facts,
- unresolved=0,
- scheduled=80,
- performed=78,
- replaced=5,
- suspended=2,
- `DAILY_STATISTICS` written,
- input unchanged=True,
- VBA preserved=True.

Fresh V87 reread confirmed 2 treatment suspensions, stored suspended=2, unresolved=0 and finalized=ΝΑΙ.

## Replacement persistence implementation

Durable placement overlays are implemented through:

- `Python/rehab_excel/daily_replacement_facts.py`
- internal `DAILY_REPLACEMENT_FACTS` sheet, `xlSheetVeryHidden`
- structured fields required to recreate `ReplacementAssignment`
- concrete-session validation
- duplicate/conflicting active target rejection
- idempotent write behavior
- reader compatibility with older workbooks without the sheet
- deterministic replacement IDs per session/date

`apply_daily_input_replacements.py` reloads persisted replacements before interactive resolution, counts them in capacity/daily state, skips already-resolved target sessions, and persists active placements into the output workbook.

`Python/tools/verify_daily_fact_reconstruction.py` reconstructs daily state from recurring/base + DAILY_INPUT + DAILY_REPLACEMENT_FACTS and compares the reconstructed statistics against saved DAILY_STATISTICS.

Focused Windows gate: **28 passed, 3 warnings in 5.14s**.

## V88 native blocker

The first reconstruction run against `ORPHAN_RESOLUTION_V88.xlsm` stopped on:

`DAILY_REPLACEMENT_FACTS row 2: invalid ReplacementDate 46313.875`

This exposed a COM timezone serialization issue for calendar-only dates. The value represents the intended 2026-10-19 date shifted to the previous UTC day at 21:00.

Fix implemented:

- numeric legacy ReplacementDate values normalize to the nearest whole Excel day,
- regression covers the exact native value `46313.875 -> 2026-10-19`,
- future ReplacementDate persistence uses ISO text (`YYYY-MM-DD`) rather than COM datetime.

Commits:

- `030acf0987d992bb2611a78f0f18df24a0b2c859`
- `7091fd91543eeaa219c370b2579712a10376651d`

Status: **reader/writer fix implemented; focused test and existing-V88 reconstruction rerun pending**.

## Immediate roadmap

1. Pull latest branch.
2. Run `python -m pytest Tests/test_daily_replacement_facts.py -q`.
3. Re-run `verify_daily_fact_reconstruction.py` against the existing V88, with expected 5 replacements, 2 suspensions, unresolved=0.
4. If reconstructed statistics exactly match stored DAILY_STATISTICS, close full authoritative daily-state restart persistence.
5. Replace snapshot ledger restore with authoritative fact removal/deactivation + full recompute.
6. Test selected/all restore with multiple simultaneous changes.
7. Resume outpatient recurring create/edit native gate.
8. Later clean Python 3.14 ZipFile warnings.

## Locked daily rules

1. recurring/base schedule for concrete date,
2. cancellations/no-shows release capacity,
3. therapist absences,
4. identify displaced existing patients,
5. apply persisted replacement overlays and redistribute existing patients into net free capacity,
6. remaining orphans resolve by placement or treatment suspension,
7. new patients only after existing displaced patients are resolved.

Operational capacity is timeslot-based. Daily facts never destructively edit recurring schedules. Recompute must be idempotent and order-independent.

## Existing native green areas

- MASTER toolbar/header UX,
- patient-centric scheduling,
- MASTER + THERAPIST_DAILY daily move sync,
- source trace,
- individual ledger restore,
- readable time axis,
- non-robotic color semantics,
- mixed orphan placement/suspension flow,
- suspension persistence across reread.

## Deferred issue

Legacy `ΑΝΑΝΕΩΣΗ` Form Control remains manual/deferred cosmetic cleanup.

## Safety constraints

- Never write directly to baseline `.xlsm` during development.
- Preview workflows use copies only.
- Preserve/verify VBA.
- Keep layout write zones narrow.
- Windows tests use `python -m pytest`.

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests
