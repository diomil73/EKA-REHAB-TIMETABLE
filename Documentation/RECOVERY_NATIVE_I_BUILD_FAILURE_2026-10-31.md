# Recovery native I build failure, 2026-10-31

## Result

The attempted fresh artifact `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_I.xlsm` was not created.

The unified builder failed during registration/operational module installation with:

`Operational module installation failed during patching operator visibility: MASTER patient status still permits strike-through`

The production source workbook hash was explicitly confirmed unchanged after the failed build.

Subsequent contract/hash/Get-Item commands failed only because the I artifact did not exist. They are not independent workbook failures.

## Diagnosis

The patient-status presentation patch itself was present and correct. The post-install verification read VBA back through Excel COM and then searched for an LF-only literal sequence. Excel COM can return CodeModule text with CRLF line endings, making the verification fail even when the installed VBA behavior is correct.

This is the same class of COM line-ending false negative seen earlier in recovery work.

## Bounded fix

- `operator_visibility_patch_vba.py` now uses `_normalize_vba()` consistently for patch input and for installed CodeModule verification.
- Both DAILY_INPUT and global-replacement code read back from COM are normalized before semantic contract checks.
- A regression test explicitly covers CRLF input and requires the post-install verification path to normalize COM text.

Commits:

- `daa515363fb07264db91f572bc8513e31664a7ea` Normalize installed VBA before visibility verification
- `6e190c31af08b032f52ccf5be7d7e5d4a85abe2d` Cover CRLF-normalized operator visibility verification

## Safety state

- Production workbook: unchanged.
- Artifact I: absent; name retired.
- No native smoke was run.
- H remains frozen evidence.
- Next artifact, only after focused source gate, is fresh `UNIFIED_CLOSE_DAY_RECOVERY_20261031_J.xlsm`.

## Next gate

Run the operator visibility/build-focused tests together with the recovery ledger/application-contract tests. If green, create J with the normal no-overwrite/source-hash safeguards.
