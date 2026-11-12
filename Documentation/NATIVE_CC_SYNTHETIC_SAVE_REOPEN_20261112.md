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
