# Project Status

Last updated: 2026-10-17

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **authoritative daily operational facts + restart-safe recomputation**, following successful native orphan resolution and suspension reread.

Latest confirmed checkpoints:

- full regression: **699 passed, 19 warnings**
- replacement load + mixed-slot gate: **31 passed in 2.04s**
- MASTER daily-overlay + ledger restore gate: **24 passed, 5 warnings in 4.39s**
- orphan suspension + statistics domain gate: **21 passed in 3.11s**
- consolidated orphan operational gate: **43 passed, 3 warnings in 5.15s**
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

The previous restart gap was durable placement overlays. Source implementation now exists:

- `Python/rehab_excel/daily_replacement_facts.py`
- internal `DAILY_REPLACEMENT_FACTS` sheet, `xlSheetVeryHidden`
- structured fields required to recreate `ReplacementAssignment`
- concrete-session validation
- duplicate/conflicting active target rejection
- idempotent write behavior
- COM-safe date serialization
- reader returns no facts for older workbooks without the sheet
- deterministic replacement IDs per session/date

`apply_daily_input_replacements.py` now reloads persisted replacements before interactive resolution, counts them in capacity/daily state, skips already-resolved target sessions, and persists active placements into the output workbook.

New full restart verifier:

- `Python/tools/verify_daily_fact_reconstruction.py`

It reconstructs daily state from recurring/base + DAILY_INPUT + DAILY_REPLACEMENT_FACTS and compares the reconstructed statistics against saved DAILY_STATISTICS.

Regressions:

- `Tests/test_daily_replacement_facts.py`
- `Tests/test_daily_fact_reconstruction_verifier.py`

Status: **implementation complete, focused source gate pending; native V88 pending**.

## Immediate roadmap

1. Run focused source tests for replacement fact storage, verifier and orphan resolution integration.
2. Generate `ORPHAN_RESOLUTION_V88.xlsm` from existing `ORPHAN_SMOKE_INPUT_V86.xlsm`, with at least one placement and one suspension.
3. Confirm replacement facts + suspension facts persist, input unchanged and VBA preserved.
4. Run full fact reconstruction verifier against V88 and require replacement/suspension counts, unresolved=0, finalized=ΝΑΙ and stored-statistics equivalence.
5. Replace snapshot ledger restore with authoritative fact removal/deactivation + full recompute.
6. Test selected/all restore with multiple simultaneous changes.
7. Resume outpatient recurring create/edit native gate.
8. Continue branch hardening and later clean Python 3.14 ZipFile warnings.

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
