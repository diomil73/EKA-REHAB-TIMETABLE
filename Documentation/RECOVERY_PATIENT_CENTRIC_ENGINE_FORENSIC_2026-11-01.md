# Recovery checkpoint: patient-centric engine forensic pivot

Date: 2026-11-01
Branch: work/history-statistics-rebuild

## Trigger

Native artifact U reproduced the same visible behavior and messages as prior S/T attempts. No meaningful operational difference was visible in MASTER or THERAPIST_DAILY, and restore still failed.

## Key forensic finding

The original patient-centric rendering engine has not been removed. Current `Python/tools/prepare_daily_input_in_place.py` still imports and calls `build_patient_centric_preview_plan(...)` from `Python/rehab_excel/patient_centric_preview.py` before composing outpatient and MASTER overlays.

The current `patient_centric_preview.py` still contains the proven replacement visualization behavior, including source-line replacement notes (`→ provider HH:MM`) and destination insertion logic.

Historical commits establish the original working lineage:
- `9f0a11433bc2dd6e6d2d692dee9ed737baa5401d` Add patient-centric timetable preview
- `1074bd1a62746dc6c532787add448f93c2f19f16` Add DAILY_INPUT reader, applier, and tests
- `e47c6d46739bfc5713ceab1067832fa205c7d9ac` Compact replacement overlay; remove source date

The older `Python/tools/apply_daily_input.py` path still applies the patient-centric `WritePlan` to a NEW workbook copy through `apply_write_plan_to_copy(...)`.

## Current interpretation

The rendering/business engine is not the primary suspect. The new in-session VBA application layer rejects the prepared plan during read-back at the first MASTER treatment presentation patch (`MASTER_SCHEDULE D8`) and rolls the transaction back. Because rollback occurs before the full plan is committed/refreshed, later THERAPIST_DAILY rendering and ledger/restore effects are not observable, creating the appearance that the entire replacement engine is broken.

## Strategy change

Stop alphabetic native rebuild iteration. Do not modify replacement solver or patient-centric renderer without evidence.

Next diagnostic must use an existing failed native artifact as INPUT and execute the older, proven copy-based `apply_daily_input.py` path to a brand-new diagnostic `.xlsm` output. This isolates:
1. patient-centric plan generation/rendering, versus
2. in-place VBA payload application/read-back/rollback.

If the copy-based output renders the expected operational changes, the patient-centric engine is exonerated and recovery work should target only the in-place VBA applier/read-back contract. If the copy-based output is also wrong, inspect the renderer/state inputs before any further VBA changes.

## Safety

- Production workbook remains immutable.
- Existing S/T/U artifacts remain frozen as evidence.
- Diagnostic output must be a fresh unique copy.
- No CloseOperationalDay, Restore All, or additional native build until this isolation test is complete.
