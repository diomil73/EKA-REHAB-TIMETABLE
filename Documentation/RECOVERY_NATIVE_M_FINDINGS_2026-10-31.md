# Native M findings and N recovery gate — 2026-10-31

## Artifact

Frozen evidence artifact:

`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_M.xlsm`

Build evidence before native testing:

- unified build completed successfully,
- production source unchanged: `True`,
- application contract: `24/24`,
- size: `470921` bytes,
- SHA-256 starts `C41AA27AEF444EE2C42D0A95C6B9EB126D65ADE55F3...`,
- last write: `2026-10-31 18:59:55` local.

M is frozen. Do not repair or reuse it.

## Native findings

The M patient-absence / restore smoke exposed three independent defects.

### 1. MASTER status reason still had red strike-through

The word/reason `διακομισμένος` remained red and struck through.

The existing normalization cleared whole-cell `Font.Strikethrough`, but the MASTER cell can contain rich-text / character-level formatting. Whole-cell normalization alone therefore does not guarantee that the reason text loses strike-through.

Recovery requirement:

- retain the expected red reason/status presentation,
- clear strike-through at both whole-cell and `Characters(...)` level,
- keep treatment strike-through behavior unchanged.

### 2. DAILY ledger UI overwrote existing THERAPIST_DAILY operational detail

The native ledger refresh changed the daily schedule and removed existing information describing the movement / replacement, including the original therapist/time and the destination therapist.

Root cause in `RefreshDailyChangeLedgerView`:

```text
startRow = last non-empty row in column A + 3
Range(A:startRow:G:startRow+200).ClearContents
```

This is unsafe because useful THERAPIST_DAILY content can continue in other columns below the last non-empty cell in column A.

This behavior is incorrect and is not an intended redesign of the daily schedule.

Recovery requirement:

- never use column A alone as the ledger anchor,
- never broadly clear an arbitrary A:G block,
- find the real last used row across the sheet,
- clear only rows previously owned by the ledger, based on ledger `DisplayRow` metadata,
- preserve all existing operational detail including `from therapist/time -> to therapist/time` information.

### 3. Restore stopped in VBA on `ApplyPayloadFileInPlace`

The VBA editor highlighted `ApplyPayloadFileInPlace` in blue during the restore attempt.

This is the confirmed immediate restore blocker.

`modDailyChangeLedger` called:

```text
ApplyPayloadFileInPlace(patchPath, "daily-change restore")
```

but that is not a canonical public procedure available to the ledger module.

The correct existing public patch engine is:

```text
ApplyValidatedDailyInputPayload(patchPath)
```

That engine already provides PATCH validation, application, semantic readback and rollback behavior and is used by the established DAILY_INPUT/replacement path.

A second deterministic compile defect was also present in the same function: `changeId` was assigned under `Option Explicit` without a local declaration. The recovery patch therefore also adds `Dim changeId As String`.

## Bounded recovery implementation

A post-install recovery patch was introduced instead of hand-editing the generated workbook or changing the solver/business logic.

Files:

- `Python/tools/patch_native_restore_recovery.py`
- `Python/tools/build_unified_preview_native_recovery.py`
- `Tests/test_native_restore_recovery_patch.py`
- `Tests/test_native_recovery_builder.py`

Relevant commits:

- `ced7628` initial bounded native recovery patch,
- `b530853` regressions for the M defects,
- `336f546` fresh native recovery unified builder,
- `5baaad3` builder integration regression,
- `99a3b10` compile-safe VBA declaration placement.

The recovery patch is exact-marker and fail-closed. If the generated VBA no longer has the expected source shape, the build must stop instead of applying a guessed transformation.

## Recovery patch contract

The generated preview must satisfy all of the following before native testing:

1. `modDailyChangeLedger` contains no `ApplyPayloadFileInPlace` reference.
2. Restore calls `ApplyValidatedDailyInputPayload(patchPath)`.
3. `ApplyDailyChangeRestorePayload` declares `changeId As String`.
4. Ledger UI no longer anchors to the last non-empty row of column A.
5. Ledger UI finds the real last used row with `Cells.Find`.
6. Ledger cleanup touches only its previously owned display rows.
7. MASTER status normalization clears both whole-cell and character-level strike-through.
8. Application contract remains 24/24.
9. Production source workbook hash remains unchanged.

## Next artifact

The next unused native artifact is:

`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_N.xlsm`

N must be built with `build_unified_preview_native_recovery.py`, not the old direct unified command.

## Native N smoke

After source tests and successful N build:

1. open only N,
2. record one patient absence,
3. verify MASTER treatment stays visible, red and struck through,
4. verify reason/status is red but not struck through,
5. verify THERAPIST_DAILY existing movement/replacement details remain unchanged,
6. verify ledger appears below the operational content,
7. restore the single ledger row,
8. verify no VBA editor / blue-highlight compile stop,
9. verify the DAILY_INPUT source row clears,
10. verify THERAPIST_DAILY returns to recurring/base state,
11. verify MASTER treatment/status return to base state,
12. verify ledger row remains visible as `Επαναφέρθηκε`,
13. verify the same workbook stays open.

`CloseOperationalDay` remains blocked until this succeeds. Therapist-absence restore remains unsupported/fail-closed.
