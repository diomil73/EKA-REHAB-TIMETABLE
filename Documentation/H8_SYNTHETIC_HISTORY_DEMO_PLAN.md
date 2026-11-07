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

## 2026-11-08 — synthetic multiweek history source gate GREEN

The operator ran the prescribed five-file synthetic-history/H9/H8 source regression gate after commits `b415d25`, `7174b19`, `41097f9`, `53ec37e`, and confirmed **30 passed in 3.70s**. Thus the deterministic cross-month fixture, in-memory history integrity checks, guarded builder and H9 pure/read-only contracts pass the requested source test selection. **This is not native Excel proof** and there has not yet been a native BM workbook build or chart/manual review.

**Authorize ONE fresh candidate:** `Excel/previews/UNIFIED_H8_DEMO_HISTORY_20261108_BM.xlsm` based on trusted frozen `Excel/previews/UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm`, with all Excel windows closed. The one-shot `Python/tools/prepare_h8_synthetic_demo.py` builds existing H8 native UI then seeds synthetic CURRENT_FINAL days plus H9 demo attendance data. Never overwrite or retry on a failed BM, and never change AZ/BJ/BK/BL or production.

Native sequence:
```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }

$source = ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm"
$demo = ".\Excel\previews\UNIFIED_H8_DEMO_HISTORY_20261108_BM.xlsm"
if (!(Test-Path -LiteralPath $source)) { throw "SAFETY STOP: AZ source missing" }
if (Test-Path -LiteralPath $demo) { throw "SAFETY STOP: BM already exists; no overwrite" }

python .\Python\tools\prepare_h8_synthetic_demo.py --source $source --output $demo
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BM native seed failed; freeze BM and do not retry" }

python .\Python\tools\audit_workbook_protection.py $demo
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BM protection audit failed; freeze BM" }

python .\Python\tools\inspect_h7_integrated_ui.py $demo
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BM toolbar audit failed; freeze BM" }
```

Expected additional `THERAPIST_ATTENDANCE_HISTORY` demo technical sheet means 24 sheets, not BL's 23. The H8 third therapist absence/working-day chart is **not integrated**; its honest unavailable-state remains. For BM native build SUCCESS expect `H8 SYNTHETIC HISTORY SEED OK`, `Persisted current-final readback: OK`, `H8 SYNTHETIC DEMO PREPARATION OK`; if any check fails, freeze with full logs and do not repeat on same BM. Only once both independent read-only audits pass may operator open the demo and verify distinct periods and visual charts. The seed is DEMO ONLY and must never be transferred to production.

## 2026-11-08 — BM synthetic seed native SAFETY STOP (operator transcript)

The operator executed the one-shot BM build after the **30 passed in 3.70s** source gate. H8 native UI build completed successfully, but **synthetic history seed failed and BM is permanently frozen diagnostic evidence**:

```
SAFETY STOP: FinalizedDaySnapshotStoreError:
FINALIZED_DAY_SNAPSHOTS row 28: SnapshotId is blank
Freeze candidate; never repeat or repair seed in same XLSM.
SAFETY STOP: DemoPreparationError: Synthetic finalized history seed failed with exit 2.
```

The operator's subsequent, independent read-only audits still returned `WORKBOOK PROTECTION AUDIT COMPLETE` and `H7 INTEGRATED AUDIT OK; workbook not changed`. They reported 24 worksheets and protected/veryHidden technical sheets, including `THERAPIST_ATTENDANCE_HISTORY` and partially added synthetic records in `FINALIZED_DAY_SNAPSHOTS` through used row 28, `THERAPIST_DAY_HISTORY` through row 183, `PATIENT_SESSION_HISTORY` through row 683, `DAY_HISTORY_EVENTS` through row 29. **These audits do NOT establish semantic validity or native demo success.** `H8 SYNTHETIC HISTORY SEED OK` and persisted-history readback did not occur. The builder must not be rerun on BM; never use BM for manual chart acceptance or production. The exact cause of partial blank ID on row 28 is **not yet verified** (possible COM range write shape / last-row offset / malformed trailing row). Avoid attributing cause without read-only evidence.

To diagnose instead of blind patching, added `Python/tools/inspect_h8_synthetic_seed_failure.py` via `bc998660`: read-only openpyxl report of actual data-row counts, missing SnapshotId rows, last six snapshot rows, and SHA256 unchanged check. It never writes or repairs the workbook and accepts only `_DEMO_` filename.

**Next operator step with all Excel windows closed:**
```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
python .\Python\tools\inspect_h8_synthetic_seed_failure.py --workbook ".\Excel\previews\UNIFIED_H8_DEMO_HISTORY_20261108_BM.xlsm"
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BM read-only diagnosis failed" }
```
Report exact diagnostic output. Then fix and regression-test generator; build **fresh previously-unused BN** only after test gate green. AZ/BJ/BK/BL remain unchanged. No retry, repair, or overwrite of BM.

## 2026-11-08 — BM read-only diagnosis confirms malformed native COM matrix write; source fix staged

Operator ran `inspect_h8_synthetic_seed_failure.py` on frozen BM, with `Workbook unchanged: True`. Exact diagnostic:
- `FINALIZED_DAY_SNAPSHOTS`: max_row 28, nonempty_rows 3 (baseline only), row 28 all six displayed cells None, missing SnapshotId.
- `THERAPIST_DAY_HISTORY`: max_row 183, nonempty_rows 33, row 183 blank SnapshotId.
- `PATIENT_SESSION_HISTORY`: max_row 683, nonempty_rows 161, row 683 blank SnapshotId.
- `DAY_HISTORY_EVENTS`: max_row 29, nonempty_rows 4, row 29 blank EventId.
- `THERAPIST_ATTENDANCE_HISTORY`: max_row 151, nonempty_rows 1, row 151 first five cells None, **sixth cell contains DEMO-H8-20260921-V1**.

These observations demonstrate that the previous nested Python tuple-of-tuples `Range.Value` COM assignment *did not produce a rectangular table*, and apparently displaced a value into a bottom-right cell while expanding UsedRange. The exact COM marshaler internals are not independently proven, but the write method is demonstrably unsafe. BM remains permanently frozen and ineligible for chart smoke; a protection/toolbar audit succeeding earlier does not overrule this fatal data-integrity finding.

Source mitigation commits:
- `ddbec123`: `Python/tools/seed_h8_synthetic_history.py` now calls `write_rows_verified()` to write **each scalar via Cells(row,col).Value2**, with record-width and nonblank identity validation, and **every cell read back immediately** before Save. Deliberately favors correctness over bulk COM speed.
- `dc41a4a`: three regression tests: exact scalar-row/column placement, dropped COM write triggers fail-fast mismatch, blank snapshot ID rejected before write. These tests replace the earlier untested bulk-write optimization; do not re-enable bulk COM without explicit native proof.

**New source gate not yet run**, no BN created. With all Excel windows closed:
```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
python -m pytest Tests/test_h8_synthetic_demo_history.py Tests/test_therapist_attendance_history_contract.py Tests/test_therapist_attendance_history_store.py Tests/test_h8_dashboard_visual_navigation.py Tests/test_h8_unique_patient_workload.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: fixed demo source tests failed" }
```
Expect 33 passes if the original 30 and three new tests all pass. Only after GREEN build a **new unused BN**, e.g. `Excel/previews/UNIFIED_H8_DEMO_HISTORY_20261108_BN.xlsm`, from trusted AZ using `prepare_h8_synthetic_demo.py` and independent protection + integrated UI audits; never reuse/overwrite BM and never mutate BL/BK/BJ/AZ. If any step fails, freeze BN and diagnose read-only. The eventual `H8 SYNTHETIC HISTORY SEED OK` must be backed by successful persisted readback, not just VBA/protection audit.

## 2026-11-08 — operator BM fix gate GREEN; pre-BN empty-COM safety enhancement

Operator reported **33 passed in 5.29s** on the five-file source regression gate after the BM scalar-COM-write mitigation, including its three new regression tests. This confirms the source gate, **not** native Excel seeding.

A final pre-native review identified one additional fail-closed mismatch: Excel `Range.Value2` can read an assigned empty string as `None`. Historical event/session rows include legitimate empty optional columns, and the earlier strict readback would incorrectly stop on the first such field. New commits:
- `3c71f01`: `write_rows_verified` permits only `None` or `""` for *explicitly expected blank optional fields*, while every nonblank value remains exactly verified after a scalar cell write.
- `d83e758`: one focused mock regression reproducing native Excel's empty-string-to-None readback behavior.

Because code changed after the **33-pass** report, do **not** infer the new test passes. **Run the source gate again before creating BN**:
```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
python -m pytest Tests/test_h8_synthetic_demo_history.py Tests/test_therapist_attendance_history_contract.py Tests/test_therapist_attendance_history_store.py Tests/test_h8_dashboard_visual_navigation.py Tests/test_h8_unique_patient_workload.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BN demo source gate failed" }
```
Expected **34 passed** if the 33 existing and one new case all pass. Only when GREEN, authorize **one fresh** `Excel/previews/UNIFIED_H8_DEMO_HISTORY_20261108_BN.xlsm` from trusted AZ with `Python/tools/prepare_h8_synthetic_demo.py`, then independent `audit_workbook_protection.py` and `inspect_h7_integrated_ui.py`. If any step fails, freeze BN and diagnose read-only; do not retry/overwrite BN. BM remains permanently frozen malformed COM-write diagnostic; BL remains accepted H8. This enhancement is source-only, no workbook was changed.

## 2026-11-08 — corrected BN source gate GREEN (34/34)

The operator ran the five-file source regression gate after the Excel COM optional-empty-cell normalization fix `3c71f01` and regression `d83e758` and confirmed **34 passed in 5.02s**. The previously confirmed **33 passed in 5.29s** was on the earlier scalar-cell write/readback version. This new operator result clears the **source-only** BN generator gate, not native Excel persistence.

**Next authoritative one-shot native test** with all Excel instances closed: `git pull --ff-only origin work/history-statistics-rebuild`; build previously nonexisting `Excel/previews/UNIFIED_H8_DEMO_HISTORY_20261108_BN.xlsm` from immutable `Excel/previews/UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm` using `Python/tools/prepare_h8_synthetic_demo.py --source ... --output ...`, then ONLY on success independently run read-only `Python/tools/audit_workbook_protection.py` and `Python/tools/inspect_h7_integrated_ui.py`. Reject output if already exists. If *any* stage fails stop, freeze BN and do not retry, reinstall or repair it. Diagnostic BM remains frozen (malformed native COM matrix write). BJ/BK/BL stay accepted/frozen. H9 production attendance capture and third H8 chart data integration remain unfinished; synthetic attendance data is illustrative only.

Expected native success lines: `H8 SYNTHETIC HISTORY SEED OK`, `Persisted current-final readback: OK`, `H8 SYNTHETIC DEMO PREPARATION OK`; 24 sheets (23 baseline + synthetic attendance history), both report sheets protected, five technical history sheets veryHidden/protected, `Workbook unchanged: True` in read-only protection audit, `H7 INTEGRATED AUDIT OK; workbook not changed`. Audits alone cannot override a failed seed. Only once both independent audits green may operator manually inspect BN synthetic daily/weekly/monthly/overall charts. No live clinical/production workbook changes.

## 2026-11-08 — BN native seed SAFETY STOP, Excel date serial coercion; no native success

Operator provided an extended console transcript. Confirmed source tests `34 passed in 3.14s`, then one-shot fresh `Excel/previews/UNIFIED_H8_DEMO_HISTORY_20261108_BN.xlsm` build. H8 VBA/UI installation succeeded, but native history seed **failed fast at the first synthetic service date**:

```
SAFETY STOP: DemoHistoryError:
FINALIZED_DAY_SNAPSHOTS: COM readback mismatch at row 4 col 2:
46286.0 != '2026-09-21'
Freeze candidate; never repeat or repair seed in same XLSM.
SAFETY STOP: DemoPreparationError: Synthetic finalized historyseed failed with exit 2.
```

Excel coerced ISO 2026-09-21 into native numeric date serial 46286.0 during `Value2` storage. The subsequent independent read-only audits still showed successful protection/toolbar state, but worksheet count **23**, original four history tables unchanged at pre-seed max rows (snapshot 3, therapist 33, patient 161, events 4), and `H7 INTEGRATED AUDIT OK`. No H9 attendance demo sheet. This explicitly proves the **native demo is NOT complete**; protection/toolbar audits cannot overrule fatal seed safety-stop. User rightly flagged the red failure output. **BN now frozen diagnostic**, just like BM; never retry or repair BN, and never open it as a demo. Trusted AZ and approved BJ/BK/BL unchanged.

Source-only remediation (native outcome still untested):
- `450588a`: in `write_rows_verified`, set each new string cell `NumberFormat = "@"` **before** writing ISO dates/times, snapshot IDs, codes, patient IDs, timestamps via scalar Excel COM `Value2`. This avoids native automatic numeric/date inference. Preserve true numeric fields.
- `fd6fe03`: regression for coercing `2026-09-21` into 46286.0 and for preserving zero-padded IDs as exact text.
- `6a7a32b`: Excel COM numeric Value2 often returns integer values as double (1.0 rather than 1). Compare expected numeric fields numerically while keeping string fields strict, and accept `None` only for expected empty optional fields.
- `4020f59`: regression for double numeric readback and for rejection of a coerced date serial when an ISO string is expected.

**New source gate pending**; no BO created. After `git pull --ff-only origin work/history-statistics-rebuild`, run:
```powershell
python -m pytest Tests/test_h8_synthetic_demo_history.py Tests/test_therapist_attendance_history_contract.py Tests/test_therapist_attendance_history_store.py Tests/test_h8_dashboard_visual_navigation.py Tests/test_h8_unique_patient_workload.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BO regression source tests failed" }
```
Expected **38** if previous 34 tests plus four additional targeted regression tests pass. Do not infer the new tests passed. Only once green create a fresh previously absent `Excel/previews/UNIFIED_H8_DEMO_HISTORY_20261108_BO.xlsm` from immutable AZ via the same guarded `prepare_h8_synthetic_demo.py`, then independent protection and toolbar audits, stopping and freezing BO on any error. BM/BN are never reused. No real production attendance integration/chart enabled. Record and prioritize the native seed result before any downstream read-only audit output.

## 2026-11-08 — BO corrected source gate GREEN (operator 38 passed in 5.11s)

Operator ran the exact five-file demo source suite after BN coercion fixes (`450588a`, `fd6fe03`, `6a7a32b`, `4020f59`) and confirmed **38 passed in 5.11s**. This verifies source-only regression coverage, including strict text formatting for ISO dates/identifiers, numeric COM double readback, and prevention of malformed BM-style matrix writes. **No BO native build has yet been reported; do not mark native success.**

**Authorize precisely one new candidate:** `Excel/previews/UNIFIED_H8_DEMO_HISTORY_20261108_BO.xlsm` from immutable trusted `Excel/previews/UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm`, with all Excel windows closed. Fail if output already exists; build and seed only once via `Python/tools/prepare_h8_synthetic_demo.py`; **only after `H8 SYNTHETIC HISTORY SEED OK`, `Persisted current-final readback: OK` and `H8 SYNTHETIC DEMO PREPARATION OK`** run independent read-only protection and integrated UI audits. If any command fails freeze BO and do not retry, repair or overwrite. A successful VBA/protection/toolbar audit does not override missing seeded-history success. Expect 24 sheets if successful. Then operator manually verifies diverse time periods and charts in this disposable test-only workbook; synthetic H9 staff attendance remains not connected to third H8 graph.

Run from repository root:
```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
$source = ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm"
$demo = ".\Excel\previews\UNIFIED_H8_DEMO_HISTORY_20261108_BO.xlsm"
if (!(Test-Path -LiteralPath $source)) { throw "SAFETY STOP: AZ source missing" }
if (Test-Path -LiteralPath $demo) { throw "SAFETY STOP: BO exists; never overwrite" }
python .\Python\tools\prepare_h8_synthetic_demo.py --source $source --output $demo
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BO native seed failed. Freeze BO; no retry." }
python .\Python\tools\audit_workbook_protection.py $demo
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BO protection audit failed. Freeze BO." }
python .\Python\tools\inspect_h7_integrated_ui.py $demo
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BO toolbar audit failed. Freeze BO." }
```
Diagnostics on previously failed BN and BM are already recorded. Both remain frozen, and accepted BL/BK/BJ and AZ must remain untouched.
