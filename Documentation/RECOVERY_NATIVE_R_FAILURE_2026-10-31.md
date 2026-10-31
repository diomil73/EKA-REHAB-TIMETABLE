# Recovery native R failure checkpoint - 2026-10-31

## Artifact
- `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_R.xlsm`
- Build succeeded with `NATIVE RECOVERY UNIFIED PREVIEW OK`.
- Production source `Rehab_Center_System_v27_1.xlsm` remained unchanged.
- Application contract verified: 24 components.
- Reported SHA-256 begins `E421BA1F90E24E7C1B148182E8D1A041ECDD7491CAF...`.
- Size: 471333 bytes.
- Last write: 2026-10-31 19:59:27 local.

## Native smoke result
Single patient absence was applied in R.

Observed MASTER behavior:
- patient treatment/absence presentation was applied;
- status/reason text `διακομισμενος` still showed red strike-through, so the status rich-text normalization remains functionally incorrect despite build-time verification.

Restore attempt:
1. User clicked the single-row restore action.
2. First message: `η ημερησια κατασταση δεν εφαρμοστηκε και οι αλλαγες αναιρεθηκαν. δοκιμστε ξανα ή επικοινωνξστε με τον διαχειριστη`.
3. After acknowledging, second message: `η επαναφορα απετυχε και οι αλλαγες αναιρεθηκαν`.
4. Restore did not complete.

This indicates the restore flow reached the canonical DAILY_INPUT patch application path, that inner DAILY_INPUT application failed and rolled back, and the outer restore transaction then also reported rollback/failure. No successful restore result was produced.

## Safety classification
- R is frozen as failed native evidence. Do not repair or retest R.
- Do not run `CloseOperationalDay`.
- Do not test therapist absence restore, selected restore, or restore-all.
- No further post-build recovery-patcher work should be done tomorrow unless required only to read evidence.

## Required immediate diagnostics before end-of-day
Preserve, if present:
- `%TEMP%\eka_daily_input_last_error.log`
- `%TEMP%\eka_daily_restore_last_error.log`

These logs are the authoritative next evidence for the nested restore failure. Do not retry before reading them.

## Tomorrow's direction
The post-build recovery patching line has consumed too many cycles and remains brittle. Tomorrow move the proven fixes into canonical source generators/modules with direct source-level regressions, then build a fresh artifact from the normal unified builder. Keep N/O/P/Q as build-tooling evidence and R as native failure evidence.
