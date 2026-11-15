<#
Runs the next H10 native call-edge verification on Windows.
Read-only Excel COM, macros/events force-disabled by the Python auditor.
No XLSM creation or mutation, no retries, and no overwriting logs.
Operator closes Excel and pulls the expected branch first.
#>
param()

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
Push-Location -LiteralPath $repo

try {
    $expectedBranch = "work/afternoon-medical-scope-20261115"
    $actualBranch = (& git branch --show-current).Trim()
    if ($LASTEXITCODE -ne 0 -or $actualBranch -ne $expectedBranch) {
        throw "SAFETY STOP: check out the approved branch before the gate"
    }

    $env:PYTHONPATH = Join-Path $repo "Python"
    $candidates = @(
        "Excel\previews\UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm",
        "Excel\previews\UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm",
        "Excel\previews\UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm"
    )
    foreach ($candidate in $candidates) {
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            throw "SAFETY STOP: frozen candidate not found: $candidate"
        }
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $focusedLog = "DIOMIL_H10_CALL_EDGE_FOCUSED_$stamp.txt"
    $fullLog = "DIOMIL_H10_CALL_EDGE_FULL_$stamp.txt"
    $nativeLog = "DIOMIL_H10_CALL_EDGE_NATIVE_$stamp.txt"
    $verifyLog = "DIOMIL_H10_CALL_EDGE_VERIFY_$stamp.txt"
    $structureLog = "DIOMIL_H10_CALL_EDGE_STRUCTURE_$stamp.txt"
    $bundle = "DIOMIL_H10_CALL_EDGE_BUNDLE_$stamp.txt"

    function Invoke-PythonLogged {
        param([string] $File, [string[]] $CommandArgs)

        if (Test-Path -LiteralPath $File) {
            throw "SAFETY STOP: refusing to overwrite log: $File"
        }
        $previous = $ErrorActionPreference
        $exitCode = 2
        try {
            $ErrorActionPreference = "Continue"
            & python @CommandArgs *> $File
            $exitCode = $LASTEXITCODE
        }
        finally {
            $ErrorActionPreference = $previous
        }
        if ($exitCode -ne 0) {
            Get-Content -LiteralPath $File -Tail 40
            throw "SAFETY STOP: failed exit=$exitCode log=$File"
        }
    }

    Invoke-PythonLogged -File $focusedLog -CommandArgs @(
        "-m", "pytest",
        "Tests/test_audit_vba_write_surfaces.py",
        "Tests/test_verify_h10_vba_profile_reports.py",
        "Tests/test_h10_native_workbook_event_hash_contract.py",
        "Tests/test_audit_integrated_workbook_structure.py",
        "-q"
    )

    Invoke-PythonLogged -File $fullLog -CommandArgs @(
        "-m", "pytest", "-ra", "--tb=short"
    )

    Invoke-PythonLogged -File $nativeLog -CommandArgs (@(
        "Python/tools/audit_vba_write_surfaces.py"
    ) + $candidates)

    $raw = Get-Content -LiteralPath $nativeLog -Raw
    if ($raw.Contains("SAFETY STOP") -or $raw.Contains("AUDIT INCOMPLETE") -or $raw.Contains("OUTPUT TRUNCATED")) {
        throw "SAFETY STOP: native report incomplete or unsafe: $nativeLog"
    }
    if (([regex]::Matches($raw, "VBA WRITE-SURFACE AUDIT COMPLETE")).Count -ne 3) {
        throw "SAFETY STOP: three completed native audits required: $nativeLog"
    }
    if (([regex]::Matches($raw, "SOURCE HASH UNCHANGED: YES")).Count -ne 3) {
        throw "SAFETY STOP: three unchanged file hashes required: $nativeLog"
    }
    if (([regex]::Matches($raw, "NO WORKBOOK WAS MODIFIED")).Count -ne 3) {
        throw "SAFETY STOP: three no-modification attestations required: $nativeLog"
    }

    Invoke-PythonLogged -File $verifyLog -CommandArgs @(
        "Python/tools/verify_h10_vba_profile_reports.py", $nativeLog
    )

    if (-not (Select-String -LiteralPath $verifyLog -Pattern "H10 VBA PROFILE EVIDENCE RECONCILIATION GREEN" -Quiet)) {
        throw "SAFETY STOP: native cross-candidate reconciliation missing: $verifyLog"
    }

    # Read-only OOXML scan checks whether Module6's VBA helpers are called
    # directly by worksheet formulas, without printing any formula text.
    Invoke-PythonLogged -File $structureLog -CommandArgs (@(
        "Python/tools/audit_integrated_workbook_structure.py"
    ) + $candidates)

    $structureRaw = Get-Content -LiteralPath $structureLog -Raw
    if ($structureRaw.Contains("SAFETY STOP") -or $structureRaw.Contains("AUDIT INCOMPLETE")) {
        throw "SAFETY STOP: structure/UDF reference scan incomplete: $structureLog"
    }
    if (([regex]::Matches($structureRaw, "READ-ONLY STRUCTURE AUDIT COMPLETE")).Count -ne 3) {
        throw "SAFETY STOP: three completed read-only structure scans required: $structureLog"
    }
    if (([regex]::Matches($structureRaw, "SOURCE HASH UNCHANGED: YES")).Count -ne 3) {
        throw "SAFETY STOP: structure scan source hashes changed: $structureLog"
    }
    if (([regex]::Matches($structureRaw, "NO WORKBOOK WAS MODIFIED")).Count -ne 3) {
        throw "SAFETY STOP: structure scan no-modification attestations missing: $structureLog"
    }
    if (([regex]::Matches($structureRaw, "VBA_UDF_FORMULA_REFERENCES:")).Count -ne 3) {
        throw "SAFETY STOP: formula UDF reference counts missing: $structureLog"
    }

    if (Test-Path -LiteralPath $bundle) {
        throw "SAFETY STOP: refusing to overwrite existing bundle"
    }
    foreach ($entry in @(
        @{ Label = "FOCUSED TESTS"; File = $focusedLog },
        @{ Label = "FULL REGRESSION"; File = $fullLog },
        @{ Label = "READ-ONLY NATIVE AUDIT"; File = $nativeLog },
        @{ Label = "CROSS-CANDIDATE VERIFIER"; File = $verifyLog },
        @{ Label = "READ-ONLY STRUCTURE AND VBA UDF FORMULA SCAN"; File = $structureLog }
    )) {
        Add-Content -LiteralPath $bundle -Value ("=== " + $entry.Label + " ===")
        Get-Content -LiteralPath $entry.File | Add-Content -LiteralPath $bundle
    }

    Write-Host "H10 CALL-EDGE GATE GREEN"
    Write-Host "UPLOAD: $bundle"
}
finally {
    Pop-Location
}
