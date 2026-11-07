# H8/H9 Synthetic Multiweek History Demo

Updated 2026-11-08. This is a **TEST-ONLY** workflow. Do not use these snapshots as clinical history or mix them into operational data.

## Objective

With only one genuine finalized example day, daily, weekly, monthly and overall filters appear similar. A deterministic simulated history can show real variations across weeks/months and exercise patient overview charts, selected-patient charts, therapist workload bands, planned vs performed care, replacement counts and safe historical aggregates.

The operator also confirmed **13 passed in 8.33s** for the separate H9 attendance-history model plus read-only loader. Those tests are GREEN but do not constitute a production roster/calendar capture feature.

## Safety boundary

All successful native evidence workbooks (**AZ, BJ, BK, BL**) remain **immutable**. The new output must be named \`Excel/previews/UNIFIED_H8_DEMO_HISTORY_20261108_BM.xlsm\`, be previously nonexistent, and be created from trusted **AZ** by the existing guarded native H8 installer. The seed script refuses a non-\`_DEMO_\` target and frozen suffixes, and refuses to insert history twice. It uses native Excel COM to preserve VBA, instead of saving an .xlsm with a non-VBA-aware writer. No source workbook, production workbook, or frozen preview is modified.

Two source files:
- \`Python/tools/prepare_h8_synthetic_demo.py\`: one-shot H8 builder + synthetic seed wrapper. No overwrite/retry.
- \`Python/tools/seed_h8_synthetic_history.py\`: deterministic fixed sample covering 2026-09-21 through 2026-10-23, weekdays only, omitting days already present in trusted source. It reuses at least 6 existing historical therapist IDs and 24 existing patient IDs; fails closed if roster is insufficient or ambiguous. It generates 1..6 scheduled treatments per therapist/day, planned/performed deviations, patient absences, cancellations, suspensions, accepted therapist replacements, and explicit same-day synthetic staff absence rows. All session/status/time/snapshot/therapist totals reconcile by construction.

It appends to the four existing historically authoritative technical tables: \`FINALIZED_DAY_SNAPSHOTS\`, \`THERAPIST_DAY_HISTORY\`, \`PATIENT_SESSION_HISTORY\`, \`DAY_HISTORY_EVENTS\`. A fifth **demo-only** sheet \`THERAPIST_ATTENDANCE_HISTORY\` follows the H9 versioned six-column provenance contract. Sheets are veryHidden and protected. No overwrite of baseline CURRENT_FINAL dates. Data carries clear \`DEMO\` snapshot IDs, synthetic reasons and \`SYNTHETIC_DEMO_APPROVED_ROSTER_NOT_REAL\` source strings. An independent read-only reopen checks the new day count, patient and therapist row counts, plus verified complete synthetic staff-attendance facts; failed candidate is frozen and never repaired.

**Important:** Existing BL/H8 third therapist attendance chart is still a truthful **no-history placeholder**. The demo attendance rows are for H9 data-model tests only; this script does not retrofit a production attendance chart or create a real source of working-day entitlement. The numerical outcomes are illustrative, not real patient or therapist measurements. Even if the source workbook has random-looking names, the seeder uses the IDs already registered in the workbook, not invented identifiers that would lose name resolution.

## Source-only gate (before creating ANY demo Excel)

Close Excel:

\`\`\`powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }

python -m pytest \`
    Tests/test_h8_synthetic_demo_history.py \`
    Tests/test_therapist_attendance_history_contract.py \`
    Tests/test_therapist_attendance_history_store.py \`
    Tests/test_h8_dashboard_visual_navigation.py \`
    Tests/test_h8_unique_patient_workload.py -q

if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Synthetic demo source tests failed" }
\`\`\`

**These source tests have NOT yet been reported run after the new generator commits.** No real demo XLSM has been created yet.

## Authorized one-shot native demo build after source gate GREEN

Close all Excel windows. Never overwrite/reuse existing output:

\`\`\`powershell
$source = ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm"
$demo = ".\Excel\previews\UNIFIED_H8_DEMO_HISTORY_20261108_BM.xlsm"
if (!(Test-Path -LiteralPath $source)) { throw "SAFETY STOP: trusted AZ not found" }
if (Test-Path -LiteralPath $demo) { throw "SAFETY STOP: BM already exists" }

python .\Python\tools\prepare_h8_synthetic_demo.py --source $source --output $demo
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BM build/seed failed. Freeze BM, no retry." }

python .\Python\tools\audit_workbook_protection.py $demo
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BM worksheet audit failed. Freeze BM." }

python .\Python\tools\inspect_h7_integrated_ui.py $demo
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BM toolbar audit failed. Freeze BM." }
\`\`\`

**Caveat:** The existing protection audit enumerates the extra fifth synthetic sheet, so expect 24 sheets instead of BL's 23. The four original technical sheets must still be protected/veryHidden; the fifth demo attendance sheet must also be protected/veryHidden. General legacy sheets remain at their existing protection baseline.

If source tests and all native read-only checks are GREEN, inspect only BM with Excel macros enabled. Check that the four report modes now differ across dates/weeks/months, charts reflect patient absence and treatment cancellations, colored therapist load reflects varied patient/day counts, and selected-person charts update. Use this exclusively as demonstration/smoke evidence; no operational day closure/amendment or production saves. Freeze on any abnormality, report exact symptoms and output.

## What remains for real H9

Actual absence-vs-working-day reporting requires verified roster, holiday/leave treatment, approved attendance provenance, finalized writer, version/correction integration and new source+native acceptance. The BM test data does not substitute for any of these.
