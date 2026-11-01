# Recovery canonical source migration - 2026-11-01

## Evidence boundary

Frozen native artifact: `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_R.xlsm`.

R build was structurally green, source hash unchanged, application contract 24. Native patient-absence smoke failed and R remains frozen evidence.

Observed native failures:

- MASTER status/reason `διακομισμενος` remained red with strike-through.
- DAILY_INPUT diagnostic:
  - phase=`verify patch`
  - patch_index=`2`
  - sheet=`MASTER_SCHEDULE`
  - cell=`D8`
  - error=`DAILY_INPUT read-back mismatch`
- restore diagnostic:
  - phase=`apply restore payload`
  - error=`Restore patch apply failed`

Read-only reproduction of R payload confirmed the failing D8 patch is:

`PATCH MASTER_SCHEDULE D8 presentation ... strike=1 ... source_tag=master_daily_non_active_overlay`

There are no text runs on the D8 patch. Therefore the failure is not a value/formula mismatch. It is the whole-cell strike read-back contract on a MASTER cell that can contain mixed/character-level formatting.

## Root cause model

Excel can return `Null` for whole-cell `Font.Strikethrough` on mixed-format cells. The previous native contract wrote/read whole-cell strike state for MASTER treatment/status presentation. That made a visually correct presentation change fail semantic read-back, and also allowed status/reason character formatting to remain struck after a whole-cell `False` assignment.

The restore path reuses the same validated DAILY_INPUT patch engine. Therefore the restore failure is downstream of the same MASTER read-back contract, not an independent solver failure.

## Canonical source changes

The temporary N-R post-build recovery path is retired. Do not use `build_unified_preview_native_recovery.py` or `patch_native_restore_recovery.py` for new artifacts. Those files remain only as forensic/history evidence until a later cleanup.

Canonical commits:

- `c98fcefc61474c963390b90f92c7a23e807bb657`
  - adds character-level MASTER strike apply/read-back to the normal operator-visibility installation path;
  - treatment non-active overlay applies and verifies strike per displayed character;
  - patient-status overlay applies/verifies no-strike per displayed character;
  - post-`ApplyMasterView` status normalization uses the same character-safe helper.
- `96ecd6ccb938aabb1630e14d166569317b5167cd`
  - moves restore recovery into canonical `daily_change_ledger_in_place_vba.py`;
  - restore PATCH rows use public `ApplyValidatedDailyInputPayload`;
  - `changeId` is explicitly declared;
  - ledger view anchors below the real last used row across THERAPIST_DAILY;
  - previous ledger content is cleared only when its own `Αλλαγές ημέρας` ownership header is present;
  - broad `A:G` clearing below column-A last row is removed.
- `0402af1ec138f08dbe7bf4130ae61c4760e28305`
  - regression coverage for character-safe MASTER strike handling.
- `fdb3770ecfdd506213c0b6e966bf744e34e867e4`
  - regression coverage for canonical restore engine and no-overlap ledger behavior.

## Next gate

Run the focused Python source gate before any new workbook build. If green, build the next fresh candidate with the normal canonical builder only:

`Python/tools/build_unified_preview.py`

Do not reuse R. Do not run CloseOperationalDay. Do not test therapist-absence restore yet.

The next native smoke remains one patient absence followed by one-row restore, verifying:

1. treatment remains visible, red and struck;
2. status/reason red but not struck;
3. THERAPIST_DAILY operational detail is preserved;
4. ledger row is displayed without overwriting operational content;
5. restore clears the exact DAILY_INPUT row and restores MASTER/THERAPIST_DAILY;
6. ledger row becomes `Επαναφέρθηκε`;
7. same workbook remains open.
