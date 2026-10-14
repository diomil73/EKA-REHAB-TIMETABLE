# Project Status

Last updated: 2026-10-14

## Current stage

The project remains on branch `feature/outpatient-scheduling` and the patient-centric multi-treatment scheduling milestone is now closed.

Latest full-suite checkpoint after cleanup: **699 passed, 19 warnings**.

The immediately preceding full-suite checkpoint before cleanup was **722 passed, 19 warnings**. The lower test count is expected because tests that exercised only retired single-treatment code were removed.

## Completed scheduling milestone

The current patient-program flow is native-verified for:

- new-patient continuation into scheduling,
- patient-centric multi-treatment selection,
- hybrid manual/automatic scheduling,
- one final atomic save,
- conflict validation with explicit override,
- exact adjacency support,
- edit-existing-program preload and resave,
- no false self-conflict from replaced persisted assignments,
- optional free-text provider/name persistence for Ρομποτικό / Εργο / Λογο / Ψυχ,
- MASTER composite rendering,
- final shell refresh in generated previews.

Retired single-treatment scheduling code has been removed while preserving the joint `auto-propose` bridge path used by the current patient-program scheduler.

## Accepted MASTER UX

Native V81 behavior is green:

- the four action buttons remain horizontally at the top,
- treatment/column headers remain visible,
- rows 1-4 are frozen during vertical scrolling,
- no UserForm/sidebar/CommandBars workaround is used.

Earlier V74-V79 persistent-control/sidebar experiments are retired.

## Outpatient architecture already in place

The branch already contains:

- `PatientType` support with inpatient/outpatient distinction,
- outpatient registry storage extensions,
- separate authoritative `OUTPATIENT_SCHEDULE`,
- unified base-schedule reads for operational logic,
- outpatient light-blue presentation rules in `THERAPIST_DAILY`,
- operational integration for daily input/replacements,
- cancellation/no-show overlays,
- outpatient schedule registration/suggestion/conflict infrastructure.

Important invariant: outpatient status changes storage/presentation routing, not operational workload/capacity semantics.

## Remaining roadmap candidates

The next milestone should be selected from the remaining outpatient/operational gaps. Highest-value candidates are:

1. Verify and, where needed, complete released cancellation-slot propagation into replacement capacity/availability calculations.
2. Native smoke-test outpatient recurring schedule create/edit flow end to end against `OUTPATIENT_SCHEDULE`.
3. Native smoke-test outpatient registration plus daily operational rendering, including light-blue presentation and source/VBA preservation.
4. Native smoke-test cancellation/no-show operational behavior and released-slot reuse.
5. Final branch/PR hardening once the remaining native operational checks are green.

## Safety constraints

- Never write directly to the baseline `.xlsm` during development or smoke testing.
- Preview workflows must operate on copied `.xlsm` files only.
- Preserve/verify VBA when rebuilding previews.
- Keep existing scheduling behavior backward-compatible unless an explicit business rule changes it.
- Windows test command is always `python -m pytest`.

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests
