# Diomil native Excel read-only smoke: isolated H10 CC copy

This is an **operator-run Windows Microsoft Excel test**. It proves read-only
native COM open/close and workbook structure, **not** actual treatment changes,
VBA execution, history persistence, patient identity, or live MASTER readiness.

## Safety constraints

- **Close all Excel windows first**.
- Never run this against the live MASTER or an authoritative CC directly.
- The CC named below is an existing frozen pilot source. It is READ only.
- The copy name `UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm` is unique.
  If it already exists, STOP, do not overwrite or delete it.
- No private patient-containing workbook may be committed to GitHub, emailed,
  or included in GitHub Actions test artifacts.
- All commands form **one** PowerShell scriptblock: failure stops the block
  without printing a fake success banner. Never resume at the next line after
  a SAFETY STOP.

From the repository root on the actual Windows/Excel PC, run once:

```powershell
& {
    $ErrorActionPreference = "Stop"

    git pull --ff-only origin work/history-statistics-rebuild
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: git pull failed" }

    $source = ".\Excel\previews\UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm"
    $copy   = ".\Excel\previews\UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm"

    if (!(Test-Path -LiteralPath $source -PathType Leaf)) {
        throw "SAFETY STOP: frozen CC source is absent"
    }
    if (Test-Path -LiteralPath $copy) {
        throw "SAFETY STOP: unique probe copy already exists"
    }
    if (!(Test-Path -LiteralPath ".\Excel\previews" -PathType Container)) {
        throw "SAFETY STOP: previews directory absent"
    }

    python .\Python\tools\prepare_native_probe_copy.py --source $source --output $copy
    if ($LASTEXITCODE -ne 0) {
        throw "SAFETY STOP: safe probe copy failed. Do not retry same path."
    }

    python .\Python\tools\probe_native_excel_readonly.py --isolated-copy $copy
    if ($LASTEXITCODE -ne 0) {
        throw "SAFETY STOP: native COM read-only check failed; freeze copy."
    }

    Write-Host "ISOLATED NATIVE EXCEL READ-ONLY PROBE GREEN" -ForegroundColor Green
}
```

This script requires Python 3.11, `pywin32` and installed Microsoft Excel on
Windows. The second command does not invoke VBA or save the workbook, and
validates the copy hash after closing Excel. The first command performs an
exclusive byte-exact copy, checks the source remained unchanged and preserves VBA.

**Expected safe output:** `ISOLATED XLSM BYTE-EXACT COPY OK`,
`NATIVE EXCEL READ-ONLY PROBE OK`, and the final green banner.

**If any failure:** stop; do not overwrite, repair, or delete the frozen copy;
report only the non-identifying terminal messages. Do not share clinical
workbooks/screenshots containing patient data. The next candidate gets a new
suffix after a source fix.

**After a green result:** this is still **NOT** proof of VBA writeback, rollback,
daily slots or import approval. A separate write-capable smoke on another
fresh, synthetic-only preview and native reopen/readback will be required before
functional demo acceptance.
