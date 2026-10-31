# Recovery Native I Pre-Build Gate — 2026-10-31

## Context

The first pre-I source gate after the in-place daily-change ledger/restore recovery reported:

- 65 passed
- 1 failed
- 2 warnings

The single failure was `Tests/test_therapist_workload_canonical_build.py::test_canonical_registration_build_installs_and_verifies_workload_strip`.

## Diagnosis

This was a stale source-inspection assertion, not a runtime or business-logic failure.

`build_registration_menu_preview._install_close_day_action_in_workbook` still installs and verifies `WORKLOAD_MODULE_NAME`, but verification was consolidated into the generic component-verification loop used for close-day, workload, global-replacement, and daily-change-ledger modules. The old test still required the previous literal error message to appear directly in source.

No workbook was built or modified because the gate stopped before the native I build.

## Fix

Commit `84fe7b177acb95730e9ac404f6a0f581bcfe3937` updates only the stale regression to verify the current semantic contract:

- workload installer call remains present;
- workload component participates in the generic verification tuple;
- the generic verification failure message remains present;
- workload component is removed during pre-clean;
- the canonical registration preview still reports the workload component.

No runtime source was changed by this fix.

## Current recovery boundary

- Production workbook remains immutable.
- Native H remains frozen as evidence.
- Native I has not been built yet.
- Daily-change restore preparation supports patient absence and session cancellation.
- Therapist-absence restore remains fail-closed until related replacement facts can be rolled back coherently.
- Restore architecture remains in-session: validate → read-only prepare → atomic native apply → semantic readback → rollback on failure → save same workbook.

## Next gate

Re-run the same focused pre-I suite. Only after it is fully green should a fresh `UNIFIED_CLOSE_DAY_RECOVERY_20261031_I.xlsm` be built.
