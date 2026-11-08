# H10 BV: Windows Excel native acceptance, safety and management-only sharing

## Status at creation

- Authorized requirements, source implementation and static marker review completed.
- **NO pytest GREEN claimed; NO native BV created or tested by assistant.**
- Frozen BO source: `Excel/previews/UNIFIED_H8_DEMO_HISTORY_20261108_BO.xlsm`.
- Frozen prior native visual proof: `Excel/previews/UNIFIED_H10_DEMO_THERAPIST_20261108_BU.xlsm`. **Never modify/reinstall BU.**
- Fresh target: `Excel/previews/UNIFIED_H10_DEMO_THERAPIST_20261108_BV.xlsm`.
- No retry/repair/overwrite on any failure. Freeze failed BV and choose new suffix after source fix.
- MANAGEMENT_SUMMARY inside BV is only an internal preview. **Never email/forward the operational XLSM**. Share ONLY the new independently generated verified .xlsx.

## Gate 1: source tests and immutable original

Run in PowerShell from `C:\Users\Dionysis\Documents\GitHub\EKA-REHAB-TIMETABLE`, all Excel closed:

```powershell
$ErrorActionPreference = "Stop"
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: git pull" }

python -m pytest `
  Tests/test_h10_planned_vs_actual_audit.py `
  Tests/test_h10_management_export.py `
  Tests/test_h10_therapist_simple_demo.py `
  Tests/test_h10_finalized_trend_export.py `
  Tests/test_h8_dashboard_visual_navigation.py `
  Tests/test_h8_unique_patient_workload.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: source tests not green" }

$bo = ".\Excel\previews\UNIFIED_H8_DEMO_HISTORY_20261108_BO.xlsm"
$bv = ".\Excel\previews\UNIFIED_H10_DEMO_THERAPIST_20261108_BV.xlsm"
if (!(Test-Path -LiteralPath $bo)) { throw "SAFETY STOP: BO missing" }
if (Test-Path -LiteralPath $bv) { throw "SAFETY STOP: BV already exists" }

python .\Python\tools\verify_h8_synthetic_history.py --workbook $bo
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BO history not verified" }

python .\Python\tools\audit_therapist_planned_vs_actual.py `
  --workbook $bo --therapist "Αργέντος"
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: planned/actual diagnostic failed" }
```

Read audit lines for the actual `date / initial / delivered / delta / replacements / approved_extra`; a positive delta alone does NOT prove a cause.

## Gate 2: single fresh native BV creation

**Only if Gate 1 fully GREEN and no Excel open:**

```powershell
python .\Python\tools\prepare_h10_therapist_simple_demo.py `
  --source $bo `
  --output $bv
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: freeze BV; no retry" }
```

Installer must report success and perform a persisted read-only reopen of VBA, protected history and MANAGEMENT_SUMMARY. Do not assume success if it printed an exception.

## Gate 3: independent read-only verification

```powershell
python .\Python\tools\verify_h8_synthetic_history.py --workbook $bv
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BV history" }

python .\Python\tools\audit_workbook_protection.py $bv
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BV protection scan" }

python .\Python\tools\inspect_h7_integrated_ui.py $bv
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BV integrated UI" }
```

A completed protection audit may contain **warnings about intentionally/unintentionally unlocked operational cells**; record warnings accurately, don't call full protection GREEN solely from exit code.

## Gate 4: human visual and interaction acceptance

Open BV once, go to therapist report. Inspect:
- Shared y-axis max for all therapists **within the SAME mode and comparable period**; selecting another row must not rescale falsely.
- Slim, short mean markers and conditional neon marker only on red filled column.
- Blue **initial plan**, red **final delivered**; positive red-over-blue dates flagged, not clipped.
- Clear horizontal gridlines/y ticks: left 0–6 by 1 where values permit; right sessions numeric step 1 when readable, more for large totals.
- Both mini trends taller and legible; third attendance chart remains unavailable without proven attendance facts.
- New **dynamic department-level appendix below the final therapist**, with four modes, numeric counts and correct ratio.
- `MANAGEMENT_SUMMARY` preview sheet, four mode buttons and aggregate-only series. It is **NOT** by itself a share-safe workbook; don't send BV.

## Gate 5: distinct whitelisted management-only export

Once gates 1–4 are acceptable, **close Excel first** and run with fresh path:

```powershell
$out = ".\Excel\previews\DIOMIL_MANAGEMENT_BV_DEMO_20261108.xlsx"
if (Test-Path -LiteralPath $out) {
    throw "SAFETY STOP: management export already exists"
}
python .\Python\tools\export_management_therapist_summary.py `
  --source $bv `
  --output $out
if ($LASTEXITCODE -ne 0) {
    throw "SAFETY STOP: never share failed management export"
}
```

The script reads authoritative finalized rows but copies **no worksheets, therapist names, patient names, technical history, VBA or external links**; it creates a fresh four-mode .xlsx with charts and validates ZIP privacy gates. Inspect output on PC before sending. **Only the new `DIOMIL_MANAGEMENT_BV_DEMO_20261108.xlsx` may be shared**, never BO/BU/BV XLSM. In current demo all source data are synthetic; production export requires a separate personal-data/privacy review.

## If any step fails

Stop. Save the text output, do not retry installer or export against the same candidate, do not repair frozen workbooks. Update `Documentation/CURRENT_HANDOFF.md` and `Documentation/PROJECT_STATUS.md` with exact result. Any next native candidate gets a new distinct suffix and full source gates.
