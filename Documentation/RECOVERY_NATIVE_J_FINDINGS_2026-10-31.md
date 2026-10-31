# Recovery native J findings — 2026-10-31

## Artifact

`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_J.xlsm`

Build evidence supplied by operator:
- unified preview build succeeded,
- production source hash remained unchanged,
- application contract verified with 24 required VBA components,
- artifact SHA-256 begins `3845F670CBB8CC8FD0C5396C8192FEE6D208204CC76...`,
- artifact size 471106 bytes,
- artifact last-write 2026-10-31 17:54:42 local.

## Native findings

Patient absence application reached the expected operational state except for two regressions:

1. MASTER status/reason text remained struck through. The patch-level source tag already writes red + `Strikethrough=False`, so the current diagnosis is that subsequent `ApplyMasterView` styling can reintroduce strike formatting after patch readback.
2. Per-row daily-change restore failed during the Python preparation phase with the user-facing generic message `Η προετοιμασία επαναφοράς απέτυχε.` No restore payload was applied, so the restore path failed closed before workbook mutation.

J is frozen as forensic evidence. Do not run `CloseOperationalDay`, do not repair J in place, and do not overwrite/rebuild the J filename.

## Bounded status fix

`operator_visibility_patch_vba.py` now installs `NormalizeMasterPatientStatusPresentation` and calls it immediately after `ApplyMasterView`. The invariant is broader and simpler: the MASTER `Κατάσταση` column must never be struck through. This avoids depending on transient formatting state from individual overlay patches.

A regression verifies the post-view normalization hook, dynamic `Κατάσταση` header lookup, and explicit `Font.Strikethrough=False` on non-empty status cells.

## Restore diagnosis and bounded fix

The exact Python backend failure from frozen J was:

`At least one non-active daily state is required`

The ACTIVE reversible change was `daily-input:2026-10-30:patient:21`. Restoring that one change removes the final non-active state and returns the date to a fully ACTIVE day. The generic patient-centric preview builder intentionally rejects an empty non-active state set, so using it directly for this restore edge case was the wrong projection path.

The bounded fix keeps the generic invariant unchanged and adds a restore-only clean-day path in `prepare_daily_change_restore_in_place.py`. When the post-restore truth has no non-active states, the preparer reconstructs only the affected operational cells from recurring/base truth and restores the corresponding MASTER treatment presentation and base patient status. No solver semantics, close/reopen behavior, or global preview invariant were changed.

Regression coverage was added for the zero-change restore branch and confirms that the clean-day path is explicit rather than weakening the generic preview builder.

## Read-only J verification after fix

The exact frozen J specimen was re-run through the restore preparer without applying the payload. The same change ID that previously failed now produced a verified atomic payload beginning with:

- `SOURCE_CLEAR DAILY_INPUT 21 A E`
- `RESTORE_CHANGE daily-input:2026-10-30:patient:21`
- `PATCH THERAPIST_DAILY B7 ... daily_change_restore_base_slot`
- `PATCH MASTER_SCHEDULE D8 ... master_daily_restore_overlay`
- `PATCH MASTER_SCHEDULE M8 value παρών ... master_daily_restore_status`

The J SHA-256 remained unchanged after the diagnostic (`J HASH UNCHANGED: True`). This proves the zero-change preparation failure is closed at the read-only backend boundary while preserving J as forensic evidence.

## Safety boundary

- production workbook remains immutable,
- J remains frozen evidence,
- no close/reopen restore architecture is reintroduced,
- therapist-absence restore remains disabled until related replacement facts can be rolled back atomically,
- no K artifact may be built unless the focused source gate is green,
- K must use a fresh filename and must not overwrite any prior preview.
