# DAILY_INPUT Recovery Dependency Checkpoint — 2026-10-30

## Context

The source-only recovery gate for the established in-place DAILY_INPUT/global-replacement preparation path stopped during pytest collection before any workbook was opened or modified.

The dependency failures encountered in order were:

1. missing `rehab_core.reclining_adjacency`,
2. missing `build_service_id_by_session` in `rehab_core.service_replacement_options`,
3. missing `rehab_excel.daily_suspension_resolution`.

These are recovery-lineage dependency gaps, not native Excel failures.

## Bounded recoveries

- `c13cbcd` recovered pure `rehab_core/reclining_adjacency.py`.
- `b899b47` recovered its regression tests.
- `e8ce13e` added only the structured `build_service_id_by_session(...)` API and required imports to the existing service-replacement module.
- `40d2bf1` added the pool-service mapping regression cases.
- `d210b95` recovered only the pure unresolved-session-to-suspension conversion API and constants.
- `7d7c541` recovered the four pure suspension regression tests.

The established `daily_suspension_resolution.py` also contained a legacy COM persistence helper that opened and closed a workbook. That persistence helper was intentionally NOT recovered. The recovery branch keeps only the pure conversion API used by the read-only global-replacement preparer. Native persistence remains the responsibility of the later in-session VBA transaction layer.

## Safety state

- Production workbook remains immutable.
- `UNIFIED_CLOSE_DAY_RECOVERY_20261030_B.xlsm` remains frozen native-smoke evidence and must not be repaired or overwritten.
- No fresh C workbook is permitted until the source-only dependency/preparation gate is green.
- The in-place `modDailyInputAction` has not yet replaced the older close/reopen implementation on the recovery branch.
- Quarantined replacement V2 patch/normalizer tooling remains excluded.

## Next gate

Run the focused source-only tests including daily suspension resolution, service mapping, adjacency, both in-place preparers, daily state and application contract. If green, recover the established in-session DAILY_INPUT VBA action and thin global-replacement wrapper, then add outpatient runtime diagnostic visibility before building a fresh uniquely named C workbook.
