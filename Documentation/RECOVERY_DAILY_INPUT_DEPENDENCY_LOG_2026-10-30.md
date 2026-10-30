# DAILY_INPUT Recovery Dependency Log — 2026-10-30

This addendum records the source-only dependency failures encountered while selectively recovering the established in-session DAILY_INPUT/global-replacement preparation path on `work/history-statistics-rebuild`.

## Safety boundary

- No production workbook was modified.
- `UNIFIED_CLOSE_DAY_RECOVERY_20261030_B.xlsm` remains frozen native-smoke evidence and is not a repair target.
- Both failures occurred during pytest collection, before any workbook/native VBA action was executed.
- Recovery remains selective. No bulk merge and no quarantined replacement V2 patch/normalizer tooling is allowed.

## Collection failure 1 — missing reclining adjacency helper

The focused read-only gate failed importing `prepare_global_replacements_in_place.py` with:

`ModuleNotFoundError: No module named 'rehab_core.reclining_adjacency'`

Inspection showed this dependency is a pure-core soft-preference helper that keeps replacement ΦΘ slots close to same-day reclining treatment. It has no Excel/COM/VBA/persistence behavior and depends only on existing recovery-core modules.

Recovered commits:

- `c13cbcd` — `Python/rehab_core/reclining_adjacency.py`
- `b899b47` — `Tests/test_reclining_adjacency.py`

## Collection failure 2 — missing structured special-service mapping API

After recovering reclining adjacency, collection advanced and then failed with:

`ImportError: cannot import name 'build_service_id_by_session' from 'rehab_core.service_replacement_options'`

The recovery branch already had `service_replacement_options.py`, including robotic and explicit special-service filtering, but lacked the later bounded API used by the established global replacement preparer to carry pool-service identity from authoritative recurring entries into materialized sessions.

Inspection confirmed `POOL_SERVICE_ID` already exists in the recovery `service_hours.py`, so no service-hours subsystem recovery was needed.

Surgical recovery:

- `e8ce13e` — add `build_service_id_by_session(...)`, `BaseScheduleEntry`/date/pattern imports, normalization helper, and structured pool mapping to the existing module.
- `40d2bf1` — add regression coverage proving pool mapping is built only from applicable authoritative provider-backed base entries.

No other replacement-option logic was replaced.

## Next gate

Re-run the source-only preparation/dependency gate including:

- `Tests/test_service_replacement_options.py`
- `Tests/test_reclining_adjacency.py`
- `Tests/test_prepare_daily_input_in_place.py`
- `Tests/test_prepare_global_replacements_in_place.py`
- `Tests/test_daily_state.py`
- `Tests/test_application_contract.py`

Only after this gate is green may the established in-session `modDailyInputAction` and thin global-replacement VBA wrapper be recovered. A fresh C workbook must not be built before those source/VBA gates are green.
