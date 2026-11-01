# Recovery native T failure — 2026-11-01

## Artifact
- `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261101_T.xlsm`
- Built by canonical `Python/tools/build_unified_preview.py`
- Source workbook remained unchanged during build.
- Application contract verified before native smoke.

## Native smoke result
The same narrow patient-absence + restore smoke was run as for S.

Observed by operator:
- visual/runtime result was effectively identical to S,
- the same user-facing failure messages appeared,
- no meaningful improvement was visible,
- restore still failed.

The exact fresh runtime error log contents have not yet been captured for T in this checkpoint. They must be read before another source change or another workbook build.

## Interpretation
- T is frozen as failed native evidence.
- Do not retest or repair T in place.
- Do not build U until the fresh T logs identify the current runtime blocker.
- Previous D9 diagnostic showed that COM `Characters()` access can fail even when whole-cell strike/color are already correct; however T proving no visible improvement means we must not assume that was the only runtime blocker.
- Status/THERAPIST_DAILY/restore observations after an aborted transaction cannot be treated as independent failures until the first failing patch in T is identified.

## Safety gate
Do not run:
- Restore all,
- therapist absence restore,
- CloseOperationalDay.
