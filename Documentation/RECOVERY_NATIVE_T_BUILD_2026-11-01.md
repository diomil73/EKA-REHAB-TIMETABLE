# Recovery native T build checkpoint — 2026-11-01

## Artifact

`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261101_T.xlsm`

## Build status

User reported the canonical T build completed successfully after commit `1b7156f` (`Fix MASTER strike read-back fallback`).

The T artifact was built only through the normal canonical builder:

`Python/tools/build_unified_preview.py`

No recovery workbook patcher or post-build VBA rewrite was used.

## Source state entering T

Focused source gate before commit:

- `60 passed`
- `3 warnings`

The warnings are the already known non-functional Python/pytest cleanup warnings and are not the native gate failure under investigation.

## Relevant fix now present

`VerifyCharacterStrike` now:

1. trusts a definitive whole-cell `Font.Strikethrough` value immediately,
2. falls back to character-level verification only when Excel returns `Null` for the whole cell,
3. fails closed if the character API itself cannot be read.

This change addresses the S diagnostic where `MASTER_SCHEDULE!D9` was already visually and semantically whole-cell struck/red, but COM `Characters()` reads raised `Member not found` and caused a false DAILY_INPUT read-back mismatch.

## Frozen evidence

`UNIFIED_CLOSE_DAY_RECOVERY_20261101_S.xlsm` remains frozen and must not be modified or retested.

## Next native gate

Run only on T:

1. one patient absence in `DAILY_INPUT`,
2. Apply,
3. verify MASTER treatment display, patient status/reason display, and preservation of existing `THERAPIST_DAILY` operational details,
4. restore only that single ledger row,
5. verify DAILY_INPUT/source row, MASTER, THERAPIST_DAILY and ledger state return correctly.

Do not run yet:

- Restore all,
- therapist absence restore,
- `CloseOperationalDay`.

Production baseline remains immutable.
