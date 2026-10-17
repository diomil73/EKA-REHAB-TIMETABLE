# Current Handoff

Last updated: 2026-10-17

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Latest native orphan preview: `Excel/previews/ORPHAN_RESOLUTION_V88.xlsm`
- V81 MASTER toolbar/header UX remains native green.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- replacement load + mixed outpatient/inpatient regression: `31 passed in 2.04s`
- MASTER daily overlay + ledger restore source gate: `24 passed, 5 warnings in 4.39s`
- orphan suspension + statistics domain gate: `21 passed in 3.11s`
- consolidated orphan operational gate: `43 passed, 3 warnings in 5.15s`
- authoritative replacement fact focused source gate: `28 passed, 3 warnings in 5.14s`
- V82 individual ledger restore: native green
- V83 MASTER active-overlay sync + restore: native green
- V84 time-axis display + restore: native green
- V85 stale-orange cleanup + restore: native green
- V87 mixed orphan resolution: native green
- V87 suspension/statistics reread persistence: native green

Python 3.14 `ZipFile.__del__` warnings remain non-blocking technical debt.

## Native V87 orphan checkpoint

Input: `ORPHAN_SMOKE_INPUT_V86.xlsm`
Output: `ORPHAN_RESOLUTION_V87.xlsm`
Date: `2026-10-19`

Confirmed:

- accepted replacements: 5
- treatment suspensions: 2
- persisted suspension facts: 2
- unresolved: 0
- scheduled: 80
- performed: 78
- replaced: 5
- suspended: 2
- source unchanged: True
- VBA preserved: True
- `DAILY_STATISTICS` created

Placements:

- ΤΣΙΚΟΒΑΣ -> Φιλιππούσης 10:45
- ΜΑΛΑΠΑΝΗ -> Πέτσιος 10:00
- ΟΛΥΜΠΙΟΣ -> Πέτσιος 10:45
- ΜΑΡΙΝΟΠΟΥΛΟΥ -> Παπακίτσου 11:30
- ΜΩΡΑΙΤΗΣ -> Αργέντος 10:45

Suspensions:

- ΠΑΠΠΟΥ: Σάρρας 09:15 -> `Αναστολή θεραπείας`
- ΣΑΒΒΑΚΗ: Σάρρας 10:45 -> `Αναστολή θεραπείας`

Fresh reread of V87 also confirmed:

- treatment suspensions reread: 2
- stored statistics treatment_suspended: 2
- stored unresolved: 0
- stored finalized: ΝΑΙ

## ACTIVE milestone: authoritative replacement facts

Source implementation is in place and the focused Windows gate is green.

New module:

- `Python/rehab_excel/daily_replacement_facts.py`

Structured very-hidden sheet:

- `DAILY_REPLACEMENT_FACTS`

Authoritative schema stores:

- ReplacementId
- TargetSessionId
- PatientId
- OriginalTherapistId
- ReplacementTherapistId
- ReplacementDate
- ReplacementTime
- Reason
- ProviderKind
- Status (`ACTIVE`)

Safety/semantics:

- validates patient/original therapist/date against the concrete target Session,
- rejects duplicate ReplacementId,
- rejects multiple conflicting active facts for the same target session/date,
- missing sheet is backward-compatible and means no persisted replacements,
- blank provider kind defaults to therapist for legacy compatibility,
- output writer is idempotent for the same semantic replacement,
- conflicting replacement for the same target/date safety-stops,
- fact sheet is `xlSheetVeryHidden`,
- recurring/base remains untouched.

`apply_daily_input_replacements.py` now:

- loads existing persisted replacement facts from the input workbook before resolution,
- includes them in capacity/state calculations,
- skips already resolved target sessions,
- uses deterministic replacement IDs: `daily-input:<session_id>:<date>`,
- persists all active replacement facts into the output workbook,
- reports loaded and newly-written replacement fact counts.

New full restart verifier:

- `Python/tools/verify_daily_fact_reconstruction.py`

It reconstructs the day from:

`recurring/base + DAILY_INPUT + DAILY_REPLACEMENT_FACTS`

then compares reconstructed scheduled/performed/replaced/absent/postponed/no-show/suspended/unresolved/finalized values against stored `DAILY_STATISTICS`.

Focused Windows result:

- `28 passed, 3 warnings in 5.14s`
- warnings are the known Python 3.14 `ZipFile.__del__` teardown issue and are non-blocking.

## V88 native reconstruction blocker and fix

Running the full reconstruction verifier against `ORPHAN_RESOLUTION_V88.xlsm` stopped on:

`DAILY_REPLACEMENT_FACTS row 2: invalid ReplacementDate 46313.875`

Root cause:

- Excel/COM persisted the calendar-only date through a timezone-shifted numeric serial.
- For the intended 2026-10-19 date, the saved serial appeared as `46313.875` (previous UTC day at 21:00).
- Simple truncation would incorrectly reconstruct 2026-10-18.

Fix implemented:

- legacy numeric `ReplacementDate` values are normalized to the nearest whole Excel day before conversion,
- exact regression added for `46313.875 -> 2026-10-19`,
- future writes store `ReplacementDate` as ISO text (`YYYY-MM-DD`) instead of COM datetime, eliminating timezone ambiguity.

Commits:

- `030acf0987d992bb2611a78f0f18df24a0b2c859` implementation
- `7091fd91543eeaa219c370b2579712a10376651d` regression

Status: **fix implemented; focused test + rerun of existing V88 verifier pending**.

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
5. load persisted replacements and redistribute existing affected patients into net free capacity,
6. resolve remaining orphans by placement or `Αναστολή θεραπείας`,
7. only then expose remaining capacity to new patients.

Recomputation must remain idempotent and order-independent.

## Exact resume point

1. Pull latest branch.
2. Run `Tests/test_daily_replacement_facts.py` focused regression.
3. Re-run `verify_daily_fact_reconstruction.py` directly against existing `ORPHAN_RESOLUTION_V88.xlsm`; no V88 rebuild is required for reader verification.
4. If green, optionally regenerate a future preview once to confirm new ISO-date writer output.
5. Then use the authoritative fact store as the foundation for ledger restore by fact removal + full recompute.
6. Test selected/all restore with multiple simultaneous changes.
7. Resume outpatient recurring create/edit native gate.
8. Clean Python 3.14 ZipFile warnings later.

## Reversible daily change ledger

Native green: active move rendering, MASTER sync, source trace, per-row restore, time-axis display, stale-orange cleanup.

Still open:

- selected/all restore with multiple simultaneous changes,
- production restore via authoritative fact removal + recompute instead of snapshot restoration.

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
