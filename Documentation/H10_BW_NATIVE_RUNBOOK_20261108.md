# H10 BW: next unique candidate after frozen BV failure

## Authoritative observed outcomes

- BV source gate: **48 passed, 1 failed**, exact source failure: `Tests/test_h10_therapist_simple_demo.py::test_simple_labels_keep_technical_cause_out_of_primary_load_chart`, expected 3 occurrences of `"Ασθενείς/ημέρα"`, actual **4**. The additional fourth occurs in the explicitly intended mini-trend series, hence test assertion updated to 4, without modifying functional label.
- Operator subsequently continued **line by line in interactive PowerShell despite a `throw`**. Avoid relying on a colored `Write-Host` completion message alone: `BV CREATED AND NATIVE INSTALL VERIFIED` was printed after a failed install and was **not evidence of success**.
- BO immutable read-only synthetic history gate **GREEN**: 25 days, 522 patients/sessions as reported, 150 therapist days, 150 attendance facts, 25 events, unchanged.
- Diagnostic `--therapist "Αργέντος"` **GREEN**, 166 finalized therapist/day report rows inspected, **one positive-delta day** `2026-09-28 | initial=4 | delivered=6 | delta=+2 | replacements=2 | approved_extra=0 | released=0 | suspensions=0`. It is compatible with 2 recorded replacements; this does not by itself prove why every transfer happened.
- BV **native installation FAILED** at `verify installed VBA and protected history sheets`, reason `H10 VBA markers missing after native reopen`. This was the first install-phase internal verifier, **not** a successful persisted reopening. The exact missing marker was not printed by the old installer.
- **Freeze BV unconditionally**, even if an .xlsm exists. DO NOT re-run, repair, delete, overwrite, save or send it. BO and BU frozen unchanged.

## Source-only fix and diagnostics

- Commit `cf49746`: installer restricted to **new BW only**. Marker readback normalized for benign VBIDE whitespace differences and now identifies **exact missing markers** plus the persisted VBA line count instead of a generic message. Early readback of imported code added. This is not proof the underlying native issue is fixed.
- Commit `d946b98`: fix regression expectation from 3 to 4 intended `"Ασθενείς/ημέρα"` usages, reserve new BW in checks, add verifier regression.
- There is **no pytest GREEN, and no BW native installation yet**. Full source tests and fresh unique-path gates still mandatory before BW.

## Single PowerShell invocation, fail-closed

Execute the **whole code block as one selection** from within repository root, all Excel closed. It wraps all commands in one scriptblock so the final success banner cannot print after a failed gate. If any command or native install fails, leave the candidate frozen and stop.

```powershell
& {
    $ErrorActionPreference = "Stop"

    git pull --ff-only origin work/history-statistics-rebuild
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull" }

    python -m pytest `
      Tests/test_h10_planned_vs_actual_audit.py `
      Tests/test_h10_management_export.py `
      Tests/test_h10_therapist_simple_demo.py `
      Tests/test_h10_finalized_trend_export.py `
      Tests/test_h8_dashboard_visual_navigation.py `
      Tests/test_h8_unique_patient_workload.py -q
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Source tests. NO BW." }

    $bo = ".\Excel\previews\UNIFIED_H8_DEMO_HISTORY_20261108_BO.xlsm"
    $bw = ".\Excel\previews\UNIFIED_H10_DEMO_THERAPIST_20261108_BW.xlsm"
    if (!(Test-Path -LiteralPath $bo)) { throw "SAFETY STOP: BO missing" }
    if (Test-Path -LiteralPath $bw) { throw "SAFETY STOP: BW already exists" }

    python .\Python\tools\verify_h8_synthetic_history.py --workbook $bo
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Frozen BO verification" }

    python .\Python\tools\prepare_h10_therapist_simple_demo.py `
      --source $bo --output $bw
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BW failed; freeze, no retry" }

    python .\Python\tools\verify_h8_synthetic_history.py --workbook $bw
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BW history" }

    python .\Python\tools\audit_workbook_protection.py $bw
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BW protection" }

    python .\Python\tools\inspect_h7_integrated_ui.py $bw
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BW integrated UI" }

    Write-Host "BW SOURCE, NATIVE AND READ-ONLY GATES PASSED" -ForegroundColor Green
}
```

If any step stops, capture the FIRST Python `SAFETY STOP:` diagnostic, not only the final PowerShell `throw`. Do not continue by selecting later lines. If all gates succeed, **only then** open BW for single visual acceptance of gridlines, true department totals, dynamic appendix and management-only preview. The management external shareable .xlsx is a **separate, opt-in export**; it has its own independent read-only proof; DO NOT share operational BO/BU/BV/BW.

## Follow-up if BW installer still fails

The new source prints `H10 VBA marker read-back mismatch: [...] ; persisted lines=N`. Record that exact list, freeze BW, and make a new suffix BX after a source fix. Never relax meaningful checks to turn a failure green.
