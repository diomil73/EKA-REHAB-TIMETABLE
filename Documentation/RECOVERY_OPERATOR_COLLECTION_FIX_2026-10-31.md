# Replacement operator collection fix — 2026-10-31

## Failure observed

The focused recovery gate stopped during pytest collection before any workbook build or native Excel smoke.

Affected imports:
- `Tests/test_replacement_operator_recovery.py`
- `Tests/test_operator_visibility_patch_vba.py`
- `Tests/test_therapist_workload_canonical_build.py`

Python reported a `SyntaxError` in `Python/rehab_excel/replacement_operator_vba.py` at the control-layout assignment line:

```python
control.Left = left: control.Top = top: control.Width = width: control.Height = height
```

This was a source-code authoring error in the Python installer. It was not a VBA runtime failure, not a solver failure, and no workbook was opened, modified, built, or repaired by the failed pytest run.

## Bounded correction

Commit `068483d0597cc39f0a37d590b09ec62daf60543c` replaced the invalid chained colon syntax with four ordinary Python assignments:

```python
control.Left = left
control.Top = top
control.Width = width
control.Height = height
```

No business logic, replacement scoring, capacity rules, persistence semantics, workbook routing, or UI contract was changed by this correction.

## Regression protection

Commit `e2d1b76f2640cf1c4205e2a436b4dabb8f29e4cf` adds an explicit installer-source regression to `Tests/test_replacement_operator_recovery.py`. The test requires the four separate Python assignments and rejects the invalid colon form.

The module import performed by the same test file also remains a collection-time syntax guard.

## Artifact safety

- No `H` workbook has been built.
- Existing recovery artifacts remain frozen.
- Production workbook remains untouched.
- Native Excel testing stays blocked until the focused Python gate is green.

## Next gate

Re-run the same focused recovery suite after pulling the two fix commits. If green, proceed to a fresh uniquely named `UNIFIED_CLOSE_DAY_RECOVERY_20261031_H.xlsm` build. Do not overwrite any prior artifact.
