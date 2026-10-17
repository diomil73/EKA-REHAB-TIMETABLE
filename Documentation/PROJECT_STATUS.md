# Project Status

Last updated: 2026-10-17

## Current stage

The project remains on branch `feature/outpatient-scheduling`.

The patient-centric multi-treatment scheduling milestone is closed. The active operational milestone is the reversible daily change ledger / dynamic daily redistribution path, with outpatient recurring create/edit still open behind it.

Latest full-suite checkpoint after cleanup: **699 passed, 19 warnings**.

Latest focused operational regression checkpoint: **31 passed in 2.04s**.

## Current native checkpoint

A real guided replacement preview now builds successfully from V81 into:

`Excel/previews/DAILY_CHANGE_LEDGER_V82.xlsm`

Stress case:

- date `2026-10-19`,
- patient `ΜΕΝΤΗΣ`,
- original `Σκαρώνης 09:15`,
- selected replacement `Μηλιδάκης 10:45`,
- candidate load shown correctly as `load 5/6 | patients 6`,
- source unchanged,
- VBA preserved,
- daily change ledger installed.

The next gate is native visual/interactive acceptance of the ledger in `THERAPIST_DAILY`, especially individual restore behavior.

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

## Daily operational rules now locked

- therapist load/capacity is measured by occupied timeslots, not raw session count,
- shared slots may contain multiple patients but consume one timeslot of daily capacity,
- replacement candidate ranking uses occupied timeslots first,
- mixed inpatient/outpatient cells must not abort otherwise valid operational movement solely because whole-cell colouring is ambiguous,
- cancellations release same-day capacity,
- therapist absences and cancellations can arrive in any order and the daily plan must be recomputable/idempotent,
- displaced existing patients take priority over new-patient allocation,
- recurring/base schedules are never destructively edited for one-day operational changes,
- date-specific moves must be reversible through a visible daily change ledger.

## Outpatient architecture already in place

The branch contains:

- `PatientType` support with inpatient/outpatient distinction,
- outpatient registry storage extensions,
- separate authoritative `OUTPATIENT_SCHEDULE`,
- unified base-schedule reads for operational logic,
- outpatient presentation rules in `THERAPIST_DAILY`,
- operational integration for daily input/replacements,
- cancellation/no-show overlays,
- outpatient schedule registration/suggestion/conflict infrastructure.

Important invariant: outpatient status changes storage/presentation routing, not operational workload/capacity semantics.

## Immediate roadmap

1. Native-smoke `DAILY_CHANGE_LEDGER_V82.xlsm` in `THERAPIST_DAILY`.
2. Verify ΜΕΝΤΗΣ appears at Μηλιδάκης 10:45 and the ledger shows the exact transition from Σκαρώνης 09:15.
3. Test individual `Επαναφορά` and verify exact visual restoration without recurring-source mutation.
4. Extend native testing to selected/all restore when multiple simultaneous preview changes are supported.
5. Replace preview snapshot restoration with authoritative overlay removal + full daily recomputation.
6. Resume outpatient recurring schedule create/edit native gate against `OUTPATIENT_SCHEDULE`.
7. Continue remaining daily operational and outpatient native checks, then branch/PR hardening.

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
