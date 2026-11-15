# Project Onboarding

Last updated: 2026-10-20

## Purpose of this file

This is the **single starting point** for a new developer, a new ChatGPT context, or anyone returning to the project after a long gap.

The old approach of discovering the project by reading several unrelated files in an uncertain order is retired. Start here first. The specialist documents listed below remain authoritative for their own domain, but this file explains what the system is, how the pieces fit together, what rules must not regress, and where to continue.

## What this product is

The project is a rehabilitation / physiotherapy scheduling and daily-operations system built on an Excel `.xlsm` workbook, but the final user experience is intentionally **application-like software**, not a normal Excel workbook.

The user should work through the MASTER screen, top-level buttons, forms and guided workflows. Worksheet tabs, raw grid navigation, internal sheet names, row numbers, internal IDs and backend mechanics are not part of the normal clinical workflow.

Excel remains the execution/storage host and the place where presentation, forms and some reporting live. Python increasingly owns scheduling/business logic because it is testable and deterministic. VBA remains only where native Excel interaction is needed and is retired incrementally rather than through a risky all-at-once rewrite.

## Canonical architecture

- **Excel / VBA**: application shell, forms, buttons, native workbook interaction, presentation and selected in-session writes.
- **Python `rehab_core`**: scheduling rules, availability, conflicts, replacements, workload, daily state, validation and optimization logic.
- **Python `rehab_excel`**: workbook adapters, render plans, bridges, VBA generators/installers and Excel-specific integration.
- **Python `tools`**: developer/verification CLIs and smoke builders.
- **Tests**: business-rule and regression contract. Development order is rule -> test -> Python implementation -> green tests -> Excel integration.

Do not move business logic into VBA merely because the UI is Excel-based.

## Workbook and source safety

- `Rehab_Center_System_v27_1.xlsm` is the immutable baseline/reference workbook.
- Development never mutates the baseline directly.
- Real workbook writes use native Microsoft Excel/COM on Windows after safety checks.
- Do not save production/staging `.xlsm` files through `openpyxl`.
- Preview/smoke workbooks are disposable copies; verify the trusted source hash remains unchanged.
- Preserve VBA and narrow write zones.
- Windows tests are always run with `python -m pytest`, never bare `pytest`.

## Authoritative scheduling model

Recurring/base schedule and daily facts are different layers.

Daily operational truth is reconstructed from:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics/views`

Daily absences, cancellations, no-shows, suspensions and replacements are overlays. They must not destructively rewrite the recurring/base programme.

Locked daily recomputation order:

1. recurring/base schedule,
2. cancellations/no-shows release capacity,
3. therapist absences,
4. identify displaced/orphan sessions,
5. apply surviving ACTIVE replacement facts,
6. resolve remaining orphans by valid placement or `Αναστολή θεραπείας`,
7. only then expose capacity to new patients.

Operational capacity is timeslot-based. Existing displaced patients have priority over new patients. `MASTER_SCHEDULE` and `THERAPIST_DAILY` must represent the same effective daily state.

Every concrete-date orphan must resolve to exactly one of:

- a valid therapist/time placement, or
- `Αναστολή θεραπείας` for that date only.

A day is not finalized while unresolved orphans remain.

## Replacement planning: current contract and open optimizer work

Replacement lifecycle/reconstruction correctness is already established, but replacement **assignment quality** is not considered complete.

When multiple therapist absences are known for the same date, all absences must be entered first and all displaced sessions must be considered together. Do not finalize replacements greedily one therapist or one patient at a time.

The future global solver must:

- maximize valid placements and minimize unresolved patients,
- protect scarce options globally,
- never let a flexible patient consume another patient's unique viable slot when a complete solution exists,
- prefer the original time unless that harms a higher-priority constraint,
- keep ΦΘ and reclining treatment adjacent/nearby where feasible and reduce unnecessary waiting,
- keep provider workload balanced; about one timeslot difference may be unavoidable, while a two-timeslot imbalance should be avoided when feasible,
- respect patient absences, existing treatments, provider availability/capability, student limits and all other hard constraints,
- treat pool and robotic operating hours as persistent configuration once implemented,
- treat 13:00 as a weak soft preference to avoid, never a forbidden time,
- use double-booking/over-capacity only as a last-resort explicit user-approved exception.

Canonical scarcity example: if patient A can use 09:00, 10:00 or 11:00 and patient B can use only 09:00, 09:00 is protected for B when a complete feasible assignment exists.

Some clinical/operational knowledge cannot reasonably become mandatory data entry, such as therapist/patient physical suitability, personality compatibility or a frail patient's tolerance for waiting. These remain selective constraints/notes or explicit user judgment. The optimizer must not invent them.

## Outpatients

Outpatients are operationally real and count normally in workload, capacity, recurring schedules, daily therapist schedules and replacement calculations.

They do not belong in the inpatient planner UI. Their recurring schedule is stored separately in authoritative `OUTPATIENT_SCHEDULE` data and merged into the common scheduling engine.

Important UI rule: `OUTPATIENT_SCHEDULE` is a hidden/protected storage sheet, **not** an end-user workspace. The user must not be asked for worksheet row numbers or internal IDs such as `outpatient:<row>`.

Create/edit should be exposed through an application-style management flow that shows business data such as patient, treatment, time, days and therapist and carries the internal target invisibly.

Outpatient visual rule in `THERAPIST_DAILY`:

- light blue only,
- never infectious/yellow,
- never robotic,
- do not inherit another patient's cell colour.

## Application UX do's and don'ts

### Do

- expose normal workflows through the MASTER/top menu, buttons, forms and guided screens,
- show business concepts, not storage mechanics,
- keep internal IDs, sheet names and row numbers invisible,
- keep technical/data sheets hidden or protected,
- make create/edit/replacement/conflict-resolution flows usable without the Excel ribbon, worksheet tabs or raw grid,
- make success messages useful: tell the user what happened and what to do or check next.

### Don't

- require worksheet navigation for ordinary work,
- ask the user to type row numbers,
- expose internal English lifecycle strings or backend implementation details,
- send the user to a hidden data sheet merely to select a record,
- design UI that only works if native Excel chrome is visible,
- turn developer smoke/debug scaffolding into the final workflow.

## Current state at 2026-10-20

Closed/green:

- restart-safe authoritative daily reconstruction,
- replacement fact lifecycle and restore correctness,
- per-row / selected / restore-all native restore correctness,
- exact daily statistics reconstruction,
- combined outpatient recurring smoke installation,
- outpatient recurring CREATE native flow.

Deferred:

- further restore performance tuning. Current low-to-mid 20-second end-user latency is accepted for now unless it regresses.

Active:

- replace the developer-only outpatient edit selector that asks for an `OUTPATIENT_SCHEDULE` row number with an application-style schedule-management/selection UI,
- then native-test same-entry edit and conflict handling,
- later implement the global multi-absence replacement optimizer described above.

## Exact resume path

For the current branch, after reading this file:

1. Read `Documentation/CURRENT_HANDOFF.md` for the exact current checkpoint and next commands/scenario.
2. Read `Documentation/DECISIONS.md` before changing business rules or UX semantics.
3. Use `Documentation/IMPLEMENTATION_MAP.md` to locate ownership in code.
4. Use `Documentation/PROJECT_STATUS.md` for milestone/risk context.
5. Use relevant focused tests as executable specification.

Do not restart work from old chat history or old smoke assumptions when these repository documents disagree.

## Documentation roles

The previous multi-file onboarding/recovery order is replaced by this file as the single front door. The specialist files are **not deleted** because they serve different living-source roles:

- `Documentation/ONBOARDING.md`: first read, system overview, non-negotiables, reading map.
- `Documentation/CURRENT_HANDOFF.md`: exact resume point, latest native/test checkpoint and immediate next steps.
- `Documentation/DECISIONS.md`: durable business, architecture and UX decisions.
- `Documentation/PROJECT_STATUS.md`: milestone state, risks, open/deferred work.
- `Documentation/IMPLEMENTATION_MAP.md`: implementation ownership and code locations.
- `CHANGELOG.md`: release/history of implementation-relevant changes.
- `UI_APPLICATION_ROADMAP.md`: detailed product/UI roadmap and historical design specification when deeper product context is needed.
- Other `Documentation/*.md`: focused contracts for a specific subsystem; consult only when working on that subsystem.

This means a newcomer no longer needs to read seven large files in sequence before understanding the project. Read **this file first**, then only the specialist document needed for the task.

## Documentation maintenance rule

Documentation is part of implementation. In the same work cycle as a meaningful change, update the appropriate specialist source and, when the change affects the project-wide mental model, update this onboarding file too.

Chat history is never the sole source of truth.
