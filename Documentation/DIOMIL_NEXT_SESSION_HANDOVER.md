# DIOMIL — NEXT SESSION HANDOVER

**Repository:** diomil73/EKA-REHAB-TIMETABLE  
**Branch:** work/afternoon-medical-scope-20261115  
**Last operator-confirmed local HEAD:** b2be9373fe1bc3695eaf6f0eb075e5f763c2e1e6  
**Last checked GitHub branch HEAD:** b2be9373fe1bc3695eaf6f0eb075e5f763c2e1e6  
**Date:** 2026-11-28

## Read first, in this order

1. Documentation/DIOMIL_MASTER_CONTINUITY_GUIDE.md — authoritative full continuity and operator instructions; continue from latest §4.253.
2. Documentation/POWERSHELL_OPERATOR_EXECUTION_STANDARD.md — mandatory safe execution method; read before giving any Windows command.
3. This file — exact current location and next action.
4. Documentation/CURRENT_HANDOFF.md — detailed technical state of daily-input verifiers and projection work.
5. Documentation/PROJECT_DELIVERY_MASTER_REPORT_20261115.md — delivery history and evidence boundaries.
6. Documentation/INTEGRATED_TEST_BASELINE_MATRIX_20261115.md — exact test/CI evidence and non-authorizing boundaries.
7. For projection details: Documentation/DAILY_INPUT_PROJECTIONS_NATIVE_READBACK_DESIGN_20261125.md.
8. For the eventual real-data DEMO: Documentation/REAL_DATA_DEMO_ACCEPTANCE_AND_FULL_HANDOFF_20261116.md and Documentation/REAL_DATA_DEMO_OPERATOR_CLI_20261116.md.

## What is already GREEN locally

- Safe Git fast-forward completed; local and origin branch were equal at b2be9373. Tracked working tree was clean. Untracked logs/backups/previews were preserved.
- PREPARED_SNAPSHOT focused test: 1 passed, exit 0.
- Targeted Python 3.14 warning diagnostic across five files: 109 passed, 1 skipped (Windows symlinks unavailable), exit 0. No ZipFile warnings appeared in the supplied output.

Do not rerun these without a specific technical reason.

## Current unresolved PC task

Identify which of the three frozen files, if any, has the required expanded MASTER K/L/M and PATIENT_PLANNER psychology headers, using read-only OOXML only:

- Excel/previews/UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm
- Excel/previews/UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm
- Excel/previews/UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm

The earlier schema comparison attempt on 2026-11-28 did not complete. The operator transcript shows split PowerShell execution: finally is not recognized, transcript remained open, and moving the active transcript failed. The printed FINAL_STATUS=B is invalid. None of the three candidates has been classified BAD by that attempt.

The earlier source Excel/previews/UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm is not suitable for the new K preview: the actual preflight correctly stopped because PATIENT_PLANNER lacks Ψυχ_Ώρα. Do not rerun its preview tool.

## Next action

Perform only the read-only schema comparison of the three frozen candidates, with baseline hash checks and exact header positions. Use the saved-script or one-paste block method in POWERSHELL_OPERATOR_EXECUTION_STANDARD.md. Do not repeat the structural auditor unless a concrete issue requires it. Do not open or modify any frozen workbook.

Only after a candidate is identified should its provenance/role be checked against the existing guide and matrix before creating a new uniquely named disposable native preview.

## Naming and status rules

New evidence filenames use descriptive purpose + timestamp + status suffix:
- _G = diagnostic completed and GOOD.
- _B = diagnostic completed and BAD.
- _P = pending/unreviewed.
- _S = wrapper failure or safety stop.

A wrapper failure is never a candidate _B.

## Guardrails

- No production workbook changes.
- No frozen candidate overwrite, repair, deletion or retry.
- No Gate 2 opening or positive operational transaction.
- Separate Python GREEN, native Excel GREEN and operational acceptance.
- Do not repeat already-green local tests without technical cause.
- Update this handover, master guide, delivery report, matrix and current handoff after every meaningful checkpoint.


## Latest engineering checkpoint — 2026-11-28

The mobile-safe DAILY_INPUT evidence work has advanced beyond the previous PROJECTIONS_REOPEN report adapter. Added:
- \`Python/rehab_excel/daily_input_pending_decision_inventory.py\`
- \`Tests/test_daily_input_pending_decision_inventory.py\`

It creates a deterministic, operation-bound snapshot of every unresolved therapist-absence target from the existing projection preflight state model, validates exact session/state identity, counts all five statuses, distinguishes original operation source hash from current evidence-workbook hash, and writes a new exclusive evidence artifact. It does not include provider options or decisions and is not a native verifier/report adapter.

Final synthetic full CI: run **46161976361**, exact code/test HEAD \`bf8c6e34c71c45208f8312ff857aafe698cff07c\`: Windows 2958 passed / 28 warnings; Ubuntu 2958 passed / 27 warnings; both safety jobs passed.

**Next mobile engineering task:** implement the separate PENDING_DECISION_INVENTORY native-readback phase-report adapter using the existing composite native source-capture diagnostic. Keep it REPORT_ADAPTER_ONLY and non-authorizing. Then continue to REPLACEMENTS_REOPEN. Do not repeat the already-green full regression absent another code change.

**PC task remains separate:** read-only schema comparison of the three frozen H10 candidates; previous split PowerShell wrapper result is invalid and no candidate was classified BAD. No frozen workbook changed.
