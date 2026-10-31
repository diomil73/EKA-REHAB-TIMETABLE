# Recovery R native failure checkpoint — 2026-10-31

## Artifact

`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_R.xlsm`

Build evidence before native smoke:
- unified build completed successfully
- production source workbook unchanged
- application contract verified: 24 components
- recovery patch verification passed
- SHA-256 prefix reported by operator: `E421BA1F90E24E7C1B148182E8D1A041ECDD7491CAF...`
- size: 471333 bytes
- last write: 2026-10-31 19:59:27 local

## Native patient-absence smoke result

R is **FAILED native evidence** and must be frozen. Do not repair or reuse R.

Observed:
1. MASTER status/reason text `διακομισμένος` remained red with strike-through. Character-level presentation is still not correct natively.
2. Restore workflow displayed:
   - `Η ημερήσια κατάσταση δεν εφαρμόστηκε και οι αλλαγές αναιρέθηκαν. Δοκιμάστε ξανά ή επικοινωνήστε με τον διαχειριστή.`
   - followed by `Η επαναφορά απέτυχε και οι αλλαγές αναιρέθηκαν.`
3. Restore did not complete.

## Diagnostics

`%TEMP%\eka_daily_input_last_error.log`

```text
timestamp=2026-10-31 20:02:22
phase=verify patch
patch_index=2
sheet=MASTER_SCHEDULE
cell=D8
error_number=-2147219201
error=DAILY_INPUT read-back mismatch
```

`%TEMP%\eka_daily_restore_last_error.log`

```text
timestamp=2026-10-31 20:03:34
phase=apply restore payload
error_number=-2147218668
error=Restore patch apply failed
```

`%TEMP%\eka_master_status_normalization_vba.log` did not exist after the successful R build, which is expected because that log is build-time verification-only.

## Root-cause direction for next session

The two native failures are linked:
- the primary concrete failure is the DAILY_INPUT validated patch engine read-back mismatch at `MASTER_SCHEDULE!D8`, patch index 2;
- restore reuses the same validated patch engine and therefore fails downstream with `Restore patch apply failed`.

Do **not** treat DAILY_INPUT apply and restore as two unrelated bugs.

The next session must start from the canonical source path, not another post-build recovery patch layer:
1. inspect the exact PATCH payload and VerifyPatch / source-tag presentation semantics for `MASTER_SCHEDULE!D8`;
2. determine why the write succeeds far enough to reach verify but read-back mismatches;
3. fix canonical source generator / verifier semantics with regression coverage;
4. separately correct native status/reason strike-through in canonical source, proving the actual rich-text behavior;
5. build a fresh artifact only after focused source tests pass.

Do not run `CloseOperationalDay` until patient/session restore UX is natively proven green.
