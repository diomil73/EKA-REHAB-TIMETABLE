# Recovery U runtime payload mismatch checkpoint

Date: 2026-11-01
Branch: `work/history-statistics-rebuild`
Artifact under investigation: `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261101_U.xlsm`

## Confirmed evidence

The latest native DAILY_INPUT failure log from U is:

- timestamp: `2026-11-01 12:37:33`
- phase: `verify patch`
- patch_index: `2`
- sheet: `MASTER_SCHEDULE`
- cell: `D8`
- error: `DAILY_INPUT read-back mismatch`

A fresh payload generated afterwards from the same U workbook using the repository tool:

`Python/tools/prepare_daily_input_in_place.py`

showed the first patches as:

1. `THERAPIST_DAILY D2`
2. `THERAPIST_DAILY H2`
3. `THERAPIST_DAILY I13`

Therefore the native runtime log and the fresh repository-generated payload do not describe the same patch ordering.

## Why this matters

This invalidates the working assumption that repeated `MASTER_SCHEDULE D8` failures necessarily reflect the current canonical payload and verifier source.

Do not apply further verifier relaxations, rich-text fixes, MASTER formatting patches, or new native build loops until the mismatch is explained.

The patient-centric/DAILY_INPUT engine itself has already been independently proven healthy through the copy-based diagnostic path (`apply_daily_input.py`) on U, with source hash unchanged and VBA preserved.

## Current hypotheses

At least one of these must be resolved before further source changes:

1. the native U run used a different payload ordering than the current repository tool now produces;
2. the runtime resolved a different script/source path than expected;
3. the workbook state used at native runtime differed materially from the state used for the fresh payload;
4. the native error log is stale relative to the workbook/payload being inspected;
5. another installed runtime path is generating/applying patches.

The resolver in canonical `modDailyInputAction` is expected to check:

- `ThisWorkbook.Path\Python\tools\prepare_daily_input_in_place.py`
- then `ThisWorkbook.Path\..\..\Python\tools\prepare_daily_input_in_place.py`

For a workbook under `Excel/previews`, the second path should resolve to the repository copy.

## Immediate next step

Enumerate every fresh payload patch with its actual one-based index, especially all `MASTER_SCHEDULE` patches, and compare the index/cell mapping against the native log.

If fresh payload `patch #2` is still not `MASTER_SCHEDULE D8`, do not change VerifyPatch. Instrument the runtime path/payload provenance first.

## Safety

- Do not modify production workbook.
- Do not rebuild U.
- Do not create V yet.
- Do not revive post-build patchers.
- Do not alter replacement engine or patient-centric renderer.
- Preserve U and the native log as forensic evidence.
