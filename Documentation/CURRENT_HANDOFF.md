# Current Handoff

Last updated: 2026-10-17

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Latest native orphan preview: `Excel/previews/ORPHAN_RESOLUTION_V87.xlsm`
- V81 MASTER toolbar/header UX remains native green.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- replacement load + mixed outpatient/inpatient regression: `31 passed in 2.04s`
- MASTER daily overlay + ledger restore source gate: `24 passed, 5 warnings in 4.39s`
- orphan suspension + statistics domain gate: `21 passed in 3.11s`
- consolidated orphan operational gate: `43 passed, 3 warnings in 5.15s`
- V82 individual ledger restore: native green
- V83 MASTER active-overlay sync + restore: native green
- V84 time-axis display + restore: native green
- V85 stale-orange cleanup + restore: native green
- V87 mixed orphan resolution: native green

Python 3.14 `ZipFile.__del__` warnings remain non-blocking technical debt.

## Native V87 orphan checkpoint

Input: `ORPHAN_SMOKE_INPUT_V86.xlsm`
Output: `ORPHAN_RESOLUTION_V87.xlsm`
Date: `2026-10-19`

Native result:

- session cancellations: 2
- accepted replacements: 5
- treatment suspensions: 2
- persisted suspension facts: 2
- unresolved sessions: 0
- scheduled: 80
- performed: 78
- replaced: 5
- suspended: 2
- source unchanged: True
- VBA preserved: True
- `DAILY_STATISTICS` created

Suspended patients:

- ΠΑΠΠΟΥ: Σάρρας 09:15 -> `Αναστολή θεραπείας`
- ΣΑΒΒΑΚΗ: Σάρρας 10:45 -> `Αναστολή θεραπείας`

Accepted placements:

- ΤΣΙΚΟΒΑΣ -> Φιλιππούσης 10:45
- ΜΑΛΑΠΑΝΗ -> Πέτσιος 10:00
- ΟΛΥΜΠΙΟΣ -> Πέτσιος 10:45
- ΜΑΡΙΝΟΠΟΥΛΟΥ -> Παπακίτσου 11:30
- ΜΩΡΑΙΤΗΣ -> Αργέντος 10:45

This confirms the mixed operational outcome works natively: one concrete day can contain both redistributions and explicit treatment suspensions, with unresolved=0.

## Important architecture distinction discovered after V87

Suspensions are now persisted authoritatively in output `DAILY_INPUT` and can be reconstructed by a fresh reader run.

The five replacement placements in V87 are **not yet persisted as authoritative reread facts**. They exist in the generated effective workbook state, but a fresh recomputation after restart cannot yet reconstruct them from durable daily facts alone.

Do not claim full daily-state restart persistence until replacement facts are persisted too.

## Persistence verifier added

New tool:

- `Python/tools/verify_orphan_resolution_persistence.py`

It re-reads `DAILY_INPUT` from the saved workbook, counts `SessionCancellationKind.TREATMENT_SUSPENDED`, reads the stored `DAILY_STATISTICS`, and can assert:

- expected suspended count,
- expected unresolved count,
- expected finalized flag.

Regression:

- `Tests/test_verify_orphan_resolution_persistence.py`

Status: implementation committed; Windows source/native verification pending.

## ACTIVE development focus

1. Verify V87 suspension reread persistence.
2. Persist replacement placements as authoritative date-specific daily facts so full daily state survives restart/recompute.
3. Then connect restore to authoritative fact removal + full recompute instead of snapshot restoration.
4. Test selected/all restore with multiple simultaneous changes.
5. Resume outpatient recurring create/edit native gate.

## Locked orphan rule

Every orphan for the concrete date must resolve as exactly one of:

- valid therapist/time placement, or
- `Αναστολή θεραπείας` for that date only.

Suspension representation:

- `DailySessionStatus.CANCELLED`
- `SessionCancellationKind.TREATMENT_SUSPENDED`
- recurring/base untouched
- original slot released
- auditable reason
- finalized day requires unresolved=0

## Locked daily recomputation order

1. recurring/base schedule for concrete date,
2. cancellations/no-shows release capacity,
3. therapist absences,
4. identify displaced/unscheduled existing patients,
5. redistribute existing affected patients into net free capacity,
6. resolve remaining orphans by placement or `Αναστολή θεραπείας`,
7. only then expose remaining capacity to new patients.

Recomputation must remain idempotent and order-independent.

## Legacy `ΑΝΑΝΕΩΣΗ`

Deferred/manual cleanup. Not a blocker. Do not reopen unless operationally relevant.

## Business rules that must not regress

- ΦΘ provider required and solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ optional free text, not solver constraints.
- Ανακλινόμενο / Πισίνα / ΕΦΑ have no provider/name field.
- Outpatients cannot receive Ρομποτικό.
- Nonblank manual constraints are never silently overwritten.
- Standard duration 45 minutes; Ανακλινόμενο 30 minutes.
- Exact adjacency is legal.
- `PATIENT_PLANNER` authoritative for inpatient recurring; `OUTPATIENT_SCHEDULE` authoritative for outpatient recurring.
- daily cancellations/absences/replacements/suspensions/redistributions are overlays, not recurring edits.
- operational capacity is timeslot-based.
- existing displaced patients precede new patients.
- MASTER and THERAPIST_DAILY must agree on effective daily state.
- layout safety zones remain narrow.

## Recovery order if context is lost

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.
