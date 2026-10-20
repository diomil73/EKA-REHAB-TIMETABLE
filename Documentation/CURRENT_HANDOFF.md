# Current Handoff

Last updated: 2026-10-20

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Trusted daily source for restore smoke copies: `Excel/previews/RERESOLVED_V90.xlsm`
- V81 MASTER toolbar/header UX remains native green.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- OOXML MASTER projection path focused gate: `59 passed in 6.85s`
- latest direct restore backend profile: total `16.509s`, pre-Excel/verification `7.976s`, Excel output `8.533s`
- latest `build_master_projection`: `0.089s` (down from `6.142s`)
- restore performance is accepted for now and deferred until end-stage polish
- parity rebuild: row/selected/all smoke copies authoritative and VBA-preserving
- per-row / selected / restore-all in-session correctness: native green
- replacement lifecycle history V87-V92: native green
- combined outpatient recurring smoke build: native green, trusted source unchanged
- outpatient recurring CREATE: native green

## CLOSED: authoritative daily truth and replacement lifecycle

Concrete daily truth remains reconstructable from:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics`

Replacement facts preserve immutable history. First placement keeps the legacy id, re-placement after restore gets `:v2`, later cycles get `:v3`, etc. Restore never rewrites a newer lifecycle fact.

## CLOSED: Excel ledger authoritative restore correctness

`DAILY_REPLACEMENT_FACTS` is authoritative. `DAILY_CHANGE_LEDGER` is a projection. The primary Excel restore path keeps the workbook open, launches Python hidden and blocking, creates rollback protection, imports the verified operational state in-session, and saves without close/reopen or visible console.

Native correctness is green for per-row, selected multi-row, and restore-all controls. User-facing `Αιτία` remains Greek-only for known replacement causes.

## DEFERRED: restore performance / UX latency

Restore latency is no longer the active engineering priority. End-user restore time around the low-to-mid 20-second range is accepted for the current development stage. Preserve the profiler and current timings for end-stage optimization.

## ACTIVE: outpatient recurring create/edit/conflict UX

Create is native green. Same-row edit persistence has been hardened in code, but the existing `EditOutpatientSchedule` InputBox that asks for an `OUTPATIENT_SCHEDULE` worksheet row is now explicitly classified as developer/test scaffolding and is **not acceptable as final UI**.

The workbook is intentionally presented as application-like software: normal users do not see worksheet tabs, the raw Excel grid, or native Excel navigation. Therefore normal workflows must not require the user to know hidden sheet names, row numbers, internal IDs or storage mechanics.

Locked UI direction:
- `OUTPATIENT_SCHEDULE` remains hidden/protected authoritative storage.
- No normal menu button should merely navigate the user to that raw storage sheet.
- Add a top-level application workflow such as `Προγράμματα εξωτερικών ασθενών` / `Διαχείριση εξωτερικών`.
- Present entries in business terms: patient, treatment, time, days, therapist.
- Provide create/edit actions from that UI; the system resolves the internal `outpatient:<row>` target itself.
- Keep `outpatient:<row>`, worksheet row numbers, internal English strings and backend mechanics out of ordinary messages/forms.
- Success messages should give useful next-step guidance or tell the user where to review the operational result, not report implementation facts.
- Developer smoke/debug tools may expose internals temporarily, but must never become the production workflow.

Current technical edit path already supports:
- stable `target_base_entry_id`,
- backend conflict validation that ignores the target row when editing,
- in-session same-row update rather than append,
- same-patient safety validation,
- removal of the old daily-view occurrence before re-rendering the edited one when relevant.

Do not ask the user to continue the native edit gate through row-number entry. First replace or wrap that developer-only selector with an application-style schedule selection flow.

## OPEN: global multi-absence replacement optimization

Before considering replacements complete, implement and validate a global assignment pass for all displaced patients after all known therapist absences for the date have been entered.

Locked requirements:
- solve all displaced patients together; do not make irreversible greedy assignments therapist-by-therapist,
- maximize valid placements and minimize unresolved patients,
- protect unique viable provider/time options from flexible patients,
- prefer original times unless this harms another patient's unique option or important reclining-treatment adjacency,
- keep ΦΘ and reclining treatment adjacent/nearby where possible and reduce patient waiting,
- treat pool/robotic operating hours as persistent configured constraints; robotic appointments can move only within those hours and must respect known provider assignments,
- respect other treatments occupying a patient's time and explicitly recorded patient absences,
- aim for balanced provider workload; a one-hour difference can be unavoidable, but avoid two-hour imbalance where feasible,
- do not demand exhaustive availability data entry for every patient,
- represent special physical suitability, therapist/patient compatibility, and waiting tolerance selectively rather than as mandatory fields,
- double-booking/over-capacity is last resort only, with a warning and explicit user approval for each choice of patient, shared treatment, and therapist,
- protect unique scarcity explicitly: if A can use 09:00/10:00/11:00 and B can use only 09:00, 09:00 belongs to B when a complete feasible assignment exists,
- avoid 13:00 only as a weak soft preference; use it when needed rather than leave a patient unresolved or violate a higher-priority rule.

Code audit findings:
- `Python/rehab_core/replacements.py::find_replacement_candidates` currently ranks single-patient candidates by workload first and exact-time availability second. This conflicts with the new global scarcity-aware requirement if used as the final decision-maker.
- `Documentation/SEQUENTIAL_THERAPIST_REPLACEMENTS.md` describes the current sequential workflow. Preserve its live workload recalculation where useful for interactive review, but add a batch/global planning phase before any accepted assignments are committed.
- `Config/provider_policy.json` and the provider-policy layer already support persistent provider profiles and date-based overrides for exclusions, blocked times, load limits, and explicit leader availability.
- The inspected `SETTINGS` schema does not currently define dedicated pool/robotic operating hours. Add a persistent service-hours configuration and include it in candidate feasibility checks.
- Do not change production workbook behavior directly. Add pure solver tests first, then preview/native verification and documentation updates.

## Exact resume point

1. Pull latest branch.
2. Do **not** continue outpatient edit testing via worksheet row-number InputBox.
3. Implement an application-style outpatient schedule management/selection UI that lists business-facing entries and internally carries the stable edit target.
4. Add pure tests confirming the UI selection maps to the intended `target_base_entry_id` without exposing worksheet rows to the user.
5. Build a fresh smoke copy and native-test: create outpatient schedule, reopen it from the management UI, edit one field, save, confirm same authoritative entry and no duplicate.
6. Native conflict gate: on a fresh scenario, attempt a known therapist/time collision and require the conflict popup/guard rather than silent double-booking.
7. Do not manually save smoke workbooks outside the tested flow. Always use a fresh copy for each independent scenario.
8. Once create/edit/conflict are native green through the application-style UI, continue to the next scheduling/UI milestone.
9. Before closing replacements, implement the global multi-absence optimizer described above.
10. Do not return to restore performance tuning yet.

## Locked orphan rule

Every orphan for the concrete date resolves as exactly one of:

- valid therapist/time placement, or
- `Αναστολή θεραπείας` for that date only.

Suspension = `DailySessionStatus.CANCELLED` + `SessionCancellationKind.TREATMENT_SUSPENDED`. Recurring/base remains untouched. A finalized day requires unresolved=0.

## Locked recomputation order

1. recurring/base schedule
2. cancellations/no-shows release capacity
3. therapist absences
4. identify displaced/orphan sessions
5. apply surviving ACTIVE replacements
6. resolve remaining orphans by placement or suspension
7. only then expose capacity to new patients

Operational capacity is timeslot-based. Existing displaced patients precede new patients. MASTER and THERAPIST_DAILY must agree.

## Documentation discipline

Meaningful changes, regressions, native results, decisions, needs, planned additions/removals/fixes must be recorded in repository documentation during the same work cycle. Chat history is not the sole source of truth.

## Legacy `ΑΝΑΝΕΩΣΗ`

Deferred/manual cosmetic cleanup. Not a blocker.

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.