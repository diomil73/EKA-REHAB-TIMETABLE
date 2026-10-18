# Project Status

Last updated: 2026-10-18

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **Restore correctness is closed green and restore performance is deferred. Active work is now outpatient recurring create/edit native verification, followed by the broader scheduling/UI roadmap.**

## Latest confirmed checkpoints

- full regression: **699 passed, 19 warnings**
- package relationship normalization + OOXML projection gate: **59 passed in 6.85s**
- latest direct restore backend profile: **16.509s total**, **7.976s pre-Excel/verification**, **8.533s Excel output**
- `build_master_projection`: **0.089s**, down from **6.142s**
- V87-V92 authoritative restore/reconstruction lifecycle: **native green**
- per-row / selected / restore-all in-session correctness: **native green**
- combined outpatient recurring smoke build: **native green**, trusted source unchanged
- outpatient recurring CREATE flow: **native green**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## CLOSED: restart-safe authoritative daily state

Operational truth is reconstructable from durable facts:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics`

V88-V92 prove restart-safe reconstruction, immutable replacement lifecycle, atomic selected/all semantics, and exact stored-statistics equivalence.

## CLOSED: Excel ledger authoritative restore correctness

`DAILY_REPLACEMENT_FACTS` is authoritative. `DAILY_CHANGE_LEDGER` is a UI/audit projection, never an independent truth source.

The normal Excel restore workflow runs the authoritative Python backend hidden and blocking, keeps the current workbook open, creates rollback protection, imports verified operational state, saves in-session, avoids visible console/close-reopen behavior, and renders Greek-only business wording in the visible `Αιτία` column.

Native correctness is green for per-row, selected, and restore-all controls.

## OPEN: replacement quality under multiple therapist absences

Replacement lifecycle correctness is green, but replacement **assignment quality** is not considered closed yet.

All known therapist absences for a date must be entered before resolution. Every patient displaced by those absences must be evaluated together in one global assignment pass. Greedy therapist-by-therapist or patient-by-patient placement is unacceptable because a flexible patient can consume the only viable slot of a constrained patient.

Requirements captured from the operational workflow:

- maximize valid placements and minimize unresolved orphans before optimizing convenience,
- protect unique candidate slots through global scarcity-aware assignment,
- prefer the original time, unless that blocks another patient's unique option or conflicts with important adjacency to a reclining treatment,
- keep physiotherapy and reclining treatments adjacent or nearby where feasible, reducing unnecessary waiting,
- respect explicit patient absences and times already occupied by the patient's other treatments,
- treat pool and robotic operating hours as persistent configuration, not as services available at every time; robotic placements must remain within operating hours and respect known provider assignments,
- balance provider workload: about one treatment-hour difference may be unavoidable, but a two-hour imbalance should be avoided when a feasible alternative exists,
- do not introduce mandatory per-patient availability data entry at this stage,
- keep patient/therapist physical suitability, compatibility, and waiting tolerance as selective case-specific knowledge rather than requiring every record to contain extensive new fields,
- permit double-booking or over-capacity only as a last-resort, explicitly warned, user-approved exception. The system must ask which remaining patient may share with whom and which therapist; it must never silently auto-approve this.

The exact numeric trade-offs among same-time preservation, reclining adjacency, and workload balance still need validation against real cases before solver weights are frozen. The business priorities and safety boundaries are now recorded in `Documentation/DECISIONS.md`.

This requirement must be addressed before the replacements topic is considered complete.

## DEFERRED: restore performance / UX latency

The trusted backend has fallen from about `45.70s` initially to `16.509s`. Current end-user restore duration in the low-to-mid 20-second range is accepted for this development stage. Further tuning is deferred until final polish unless a regression appears.

## ACTIVE: outpatient recurring create/edit

The functional path supports recurring outpatient creation, stable-row editing via `target_base_entry_id`, seven suggestion modes, therapist double-booking checks, and the conflict popup. Combined smoke installation is native green and the CREATE flow has been confirmed native green. Remaining native gates are edit-same-row and conflict behavior.

User-facing success messages should be operationally useful: after successful actions they should guide the user to the relevant next step or verification view, rather than report implementation details such as whether the workbook stayed open.

## Immediate roadmap

1. Finish outpatient recurring edit verification: same `outpatient:<row>` must be updated with no duplicate row.
2. Finish outpatient conflict verification: known therapist/time overlap must surface the conflict guard rather than silently saving.
3. Keep smoke scenarios fresh and do not manually save outside the tested flow.
4. Before closing replacements, design and implement global multi-absence replacement optimization with scarcity-aware assignment and explicit hard/soft constraints.
5. Continue directly to the next scheduling/UI milestone after these gates.
6. Revisit restore performance only near completion.

## Locked daily rules

1. recurring/base schedule
2. cancellations/no-shows release slots
3. therapist absences
4. identify displaced/orphan sessions
5. apply surviving ACTIVE replacements
6. unresolved orphans resolve through placement or treatment suspension
7. only then expose capacity to new patients

Operational capacity is timeslot-based. Daily facts never destructively edit recurring schedules. Existing displaced patients precede new patients. MASTER and THERAPIST_DAILY must agree.

## Safety constraints

- never write directly to baseline `.xlsm` during development
- preview workflows use copies
- preserve and verify VBA
- keep layout write zones narrow
- existing replacement facts must never be re-appended during rerun
- partial previews created before later failure must be removed
- bulk restore validates the whole set before any mutation
- Excel ledger restore mutates authoritative fact lifecycle, never old snapshot cells directly
- `DAILY_CHANGE_LEDGER` remains a projection of authoritative facts
- normal ledger UI must not close/reopen the workbook or expose a console window
- normal ledger UI must not expose internal English reason labels in `Αιτία`
- Windows tests use `python -m pytest`

## Documentation discipline

Every meaningful evolution must be recorded in repository documentation during the same work cycle. Chat history is not the sole source of truth.

## Deferred issue

Legacy `ΑΝΑΝΕΩΣΗ` Form Control remains manual/deferred cosmetic cleanup.

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests
