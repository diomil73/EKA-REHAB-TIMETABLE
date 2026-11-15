# H10 VBA project and UserForm inventory: read-only runbook

Date: 2026-11-15  
Branch: work/afternoon-medical-scope-20261115

## Purpose

After the complete four-candidate structural audit, inventory the distinct VBA projects and UserForms without extracting or printing VBA source code and without changing any workbook.

Tool:
- Python/tools/audit_vba_project_inventory.py
- Tests/test_audit_vba_project_inventory.py

The tool:
- opens each XLSM in a separate native Excel process;
- sets Excel automation security to ForceDisable before opening the workbook;
- disables events and screen updates and sets calculation to manual;
- opens with ReadOnly=True, UpdateLinks=0, AddToMru=False;
- reports VBA component name/type/code-line count and UserForm control names/types only;
- does not print VBA source text, cell values, formulas, captions or patient/staff information;
- verifies the workbook SHA-256 and byte length before and after;
- deduplicates byte-identical candidates, preserving the original candidate's audit status;
- closes without saving.

It requires Windows, Microsoft Excel, pywin32, and permission to access the VBA project object model. If access is blocked or the VBA project is locked, the tool returns nonzero. Do not change Excel Trust Center / AccessVBOM settings to force the audit. Report the exact stop instead.

## Candidate facts from the completed structural audit

Full report: DIOMIL_H10_STRUCTURE_20261115_165132.txt.

| Candidate | Workbook SHA-256 | Sheets | VBA bytes | VBA SHA-256 | Distinctive sheet |
|---|---|---:|---:|---|---|
| UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm | 4df997b6a5e5b0b56b37127ee00dd8710cc26c7c118f1742d58c9969bcbe7afa | 26 | 920,576 | 8613575ee50826394f336e4f2fd864579f087a2b1e2da4fc9883582e54a1e800 | RECURRING_ASSIGNMENT_IDS, A1:C180 |
| UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm | fa9765da7079137bcb1ff935c3d08bc84b409c8335f9cd719f3583b89e4fc67d | 26 | 921,088 | 7742d872c953f2b026f66097f5a61df84ef18ac151b523e88e901f10bc5a0193 | DIOMIL_NATIVE_PROBE, A1 |
| UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm | 1bd5e58e905f91acb55e5a706cd8b4563b5a98c170a9ec60fe76c95617420b13 | 25 | 1,305,600 | 285b7f4aa20a195535d300928f26b2884255797c1817d4f2a7e32ac82b4033b9 | no additional sheet |
| UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm | same as preceding row | 25 | 1,305,600 | same as preceding row | byte-identical duplicate |

All four have 5 defined names, 11 drawings, 2 charts, 2 VML drawings, and no external links/connections/query tables/pivots/slicers/ActiveX/embeddings/custom UI in this package inventory. The third and fourth files have the same whole-file SHA-256, not merely the same VBA project; audit the component inventory only once for that pair.

Common formula/cache observations include:
- PATIENT_PLANNER: 2,745 formulas, 465 cached results.
- MASTER_SCHEDULE: 1,021 formulas, 549 cached results.
- STATISTICS: 43 formulas, 43 cached results.
- PATIENTS: 549 formulas, 93 cached results.
- OUTPATIENT_SCHEDULE: header row only.
- PATIENT_HISTORY_REPORT: empty.
- History sheets include 523 patient-session rows, 151 therapist-day rows, 151 therapist-attendance-history rows, 26 finalized-day snapshots and 26 day-history events.

The low cached-result counts are a diagnostic observation only, not proof of corrupt formulas. The _PY_LISTS role hint is heuristic and is not proof that the lookup sheet stores patient records.

## Frozen-artifact and provenance boundary

The authoritative 2026-11-12 handoff identifies the CC workbook as the trusted synthetic source and records:
- CC read-only native probe A as frozen;
- synthetic native Save/reopen A as frozen;
- recurring AssignmentID migration A as frozen, with 179 immutable IDs independently verified.

These are evidence/reference artifacts only. Never edit, save, overwrite, reinstall, or reuse them as writable candidates. The 179-ID migration proves that bounded synthetic ID migration scenario; it does not prove native permanent-event append/rollback or certify every patient/staff-bearing sheet as a fully synthetic whole-workbook fixture.

The full synthetic integration workbook remains NOT CERTIFIED. The next VBA inventory is read-only and does not select a writable baseline.

## Operator gate

Close all existing Excel windows first. Pull the branch, run the focused tests, then full regression. Only if both are green, run the read-only VBA inventory for the four manifest entries. The fourth file is automatically skipped as byte-identical to the third.

```powershell
$ErrorActionPreference = "Stop"

git pull --ff-only origin work/afternoon-medical-scope-20261115
if ($LASTEXITCODE -ne 0) {
    throw "SAFETY STOP: Git pull failed"
}

$env:PYTHONPATH = (Join-Path (Get-Location).Path "Python")

# Focused tests for both read-only auditors.
$focusedLog = Join-Path $PWD (
    "DIOMIL_FOCUSED_AUDIT_TESTS_{0}.txt" -f (Get-Date -Format "yyyyMMdd_HHmmss")
)
$previousEAP = $ErrorActionPreference
try {
    $ErrorActionPreference = "Continue"
    python -m pytest `
      Tests/test_audit_integrated_workbook_structure.py `
      Tests/test_audit_vba_project_inventory.py -q *> $focusedLog
    $focusedCode = $LASTEXITCODE
}
finally {
    $ErrorActionPreference = $previousEAP
}
"EXIT CODE: $focusedCode" | Add-Content -LiteralPath $focusedLog
Get-Content -LiteralPath $focusedLog -Tail 18
if ($focusedCode -ne 0) {
    Write-Host "FOCUSED TEST REPORT: $focusedLog"
    throw "SAFETY STOP: Focused auditor tests failed"
}

# Full source regression after the new VBA inventory code.
$fullLog = Join-Path $PWD (
    "DIOMIL_FULL_REGRESSION_{0}.txt" -f (Get-Date -Format "yyyyMMdd_HHmmss")
)
$previousEAP = $ErrorActionPreference
try {
    $ErrorActionPreference = "Continue"
    python -m pytest -ra --tb=long *> $fullLog
    $fullCode = $LASTEXITCODE
}
finally {
    $ErrorActionPreference = $previousEAP
}
"EXIT CODE: $fullCode" | Add-Content -LiteralPath $fullLog
Get-Content -LiteralPath $fullLog -Tail 18
if ($fullCode -ne 0) {
    Write-Host "FULL REGRESSION REPORT: $fullLog"
    throw "SAFETY STOP: Full regression failed"
}

# Native Excel inspection only. No macros run and no workbook is saved.
$vbaLog = Join-Path $PWD (
    "DIOMIL_H10_VBA_INVENTORY_{0}.txt" -f (Get-Date -Format "yyyyMMdd_HHmmss")
)
$previousEAP = $ErrorActionPreference
try {
    $ErrorActionPreference = "Continue"
    python Python\tools\audit_vba_project_inventory.py `
      "Excel\previews\UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm" `
      "Excel\previews\UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm" `
      "Excel\previews\UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm" `
      "Excel\previews\UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm" *> $vbaLog
    $vbaCode = $LASTEXITCODE
}
finally {
    $ErrorActionPreference = $previousEAP
}
"PYTHON EXIT CODE: $vbaCode" | Add-Content -LiteralPath $vbaLog
Get-Content -LiteralPath $vbaLog -Tail 45
Write-Host "VBA INVENTORY REPORT: $vbaLog"
Write-Host "VBA INVENTORY EXIT CODE: $vbaCode"
if ($vbaCode -ne 0) {
    throw "SAFETY STOP: VBA inventory incomplete; upload the full TXT"
}
```

## Expected result

For each distinct VBA project:
- component count;
- each component's name, type and code line count;
- each UserForm's direct control count and control names/types;
- macros disabled, events disabled, calculation manual, workbook read-only;
- source hash unchanged;
- NATIVE VBA PROJECT INVENTORY COMPLETE.

The third/fourth pair should show a byte-identical duplicate skip. If the VBA project object model is inaccessible, or a form's controls cannot be enumerated, stop and upload the full report. Do not enable Trust Center access, unlock a VBA project, edit a workbook, or substitute another candidate.

## Next engineering step after successful inventory

Compare the component inventories across the three distinct VBA projects. Then map existing VBA write surfaces, UserForms and Python installers to the acceptance matrix. Only after this mapping should a new, unique, fully synthetic disposable workbook be designed/built. The frozen H10 files remain immutable and are not the final writable test workbook.
