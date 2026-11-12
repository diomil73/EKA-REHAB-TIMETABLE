# Native Excel write/save/reopen A: new synthetic H10 CC copy only

**Source proof completed by operator (2026-11-12):**
`ISOLATED XLSM BYTE-EXACT COPY OK`; `NATIVE EXCEL READ-ONLY PROBE OK`;
`NATIVE EXCEL READ-ONLY PROBE GREEN`. The immutable CC and previous
read-only probe A were unchanged.

**Purpose of the NEXT run:** one genuine Excel COM `Save`, close,
read-only reopen and marker readback in a NEW, isolated `_DEMO_` copy.
Do not use patient grid cells or real patients: the only intended addition
is `DIOMIL_NATIVE_PROBE!A1`.

All Excel sessions closed. Run the entire single block in PowerShell
from the local repository root:

```powershell
& {
    $ErrorActionPreference = "Stop"
    git pull --ff-only origin work/history-statistics-rebuild
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: git pull" }

    $cc = ".\Excel\previews\UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm"
    $out = ".\Excel\previews\UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm"
    if (!(Test-Path -LiteralPath $cc -PathType Leaf)) {
        throw "SAFETY STOP: synthetic CC not present"
    }
    if (Test-Path -LiteralPath $out) {
        throw "SAFETY STOP: write candidate A already exists; do not overwrite"
    }

    python .\Python\tools\verify_h8_synthetic_history.py --workbook $cc
    if ($LASTEXITCODE -ne 0) {
        throw "SAFETY STOP: synthetic source history failed; no candidate"
    }

    python .\Python\tools\prove_synthetic_cc_native_write.py --source $cc --output $out
    if ($LASTEXITCODE -ne 0) {
        throw "SAFETY STOP: freeze failed native candidate; no retry"
    }

    python .\Python\tools\verify_h8_synthetic_history.py --workbook $out
    if ($LASTEXITCODE -ne 0) {
        throw "SAFETY STOP: synthetic history changed or readback failed"
    }

    Write-Host "SYNTHETIC CC NATIVE SAVE, REOPEN AND HISTORY GREEN" -ForegroundColor Green
}
```

**No retry using the same filename after any failure**; preserve the
failure evidence and do not run later commands out of context.
The new candidate is not a working operational workbook. No real
patient or therapist data is imported. Do not upload the XLSM.
The successful read-only probe A remains frozen and untouched.

This still does **not** prove treatment-change VBA Apply, Restore,
authoritative commit, user-session protection or finalized history.
Those require separate, explicit native Excel acceptance. The purpose
here is to validate the new diagnostic-sheet save/reopen mechanism only.


## Operator-confirmed native result — 2026-11-12

The operator ran the entire prescribed PowerShell block on local Windows
with real installed Microsoft Excel, using the new, isolated synthetic H10
candidate A. The pasted terminal output reported, in order:

- `H8 SYNTHETIC HISTORY CROSS-SHEET VERIFICATION OK` before write:
  25 days, 522 patient-session facts, 150 therapist-day facts,
  150 attendance facts, 25 finalization events, source unchanged.
- `SYNTHETIC NATIVE EXCEL SAVE REOPEN GREEN`, diagnostic sheet read-back YES,
  VBA preserved YES, original synthetic CC unchanged YES.
- `H8 SYNTHETIC HISTORY CROSS-SHEET VERIFICATION OK` after write:
  identical five aggregate counts, output workbook unchanged by verification.
- Final runbook completion banner
  `SYNTHETIC CC NATIVE SAVE, REOPEN AND HISTORY GREEN`.

**Disposition:** mark `UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm`
as frozen successful native write/reopen evidence. Never repurpose,
overwrite, or repeat the test using this destination. The earlier read-only
probe copy `UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm`
also remains frozen successful evidence. The CC source itself is unchanged.

**Exact scope:** proves a new diagnostic worksheet's native save and
persisted read-back, macro preservation and consistency of preexisting
synthetic H8 history. It does NOT prove daily Apply, Restore, clinical
recurrence changes, therapist substitutions, in-session protection, or
authoritative finalization after such changes. Those require a distinct,
new synthetic artifact and verified operator acceptance.


## CURRENT SAFETY NOTE (2026-11-12)

This historical native command has **already completed** on the operator's
Windows machine (see `Documentation/AUTHORITATIVE_RESUME_20261112.md`).
Its output A is accepted and FROZEN. **DO NOT RUN THIS SCRIPT AGAIN** against
the same output filename, do not overwrite/repair/delete it, and do not treat
this proof as permission to modify MASTER, CC, clinical data or permanent
therapy event history.

To resume development, start with `Documentation/ONBOARDING.md` and
`Documentation/AUTHORITATIVE_RESUME_20261112.md`. For a NEW native scenario
first implement source/tests, select a never-before-used synthetic candidate,
use one complete fail-closed PowerShell scriptblock and independently verify
history, VBA, source hashes and actual persisted semantics. On any failure
freeze that candidate and record the exact first safety error; do not retry
the same artifact.
