# Worksheet protection: first source-derived write map | 2026-11-04

## Current verified baseline

BF: `Excel/previews/UNIFIED_PATIENT_HISTORY_H7_20261104_BF.xlsm`.
The read-only inventory completed with 23 sheets and no changes. `ProtectStructure=False`, `ProtectWindows=False`. 18 sheets had `ProtectContents=False`; five were already protected:
- `PATIENT_HISTORY_REPORT`
- `FINALIZED_DAY_SNAPSHOTS`
- `THERAPIST_DAY_HISTORY`
- `PATIENT_SESSION_HISTORY`
- `DAY_HISTORY_EVENTS`

All UsedRange.Locked values reported True. This is not effective edit prevention while ProtectContents=False. The four technical history sheets are veryHidden; patient report is visible. Therapist report was unprotected in the baseline audit, but its renderer explicitly Protects after rendering; verify again after VBA/source audit before any change.

## Confirmed write paths from tracked source

This is a source-derived map, not yet an exhaustive workbook-VBA audit. Dynamic/untracked workbook modules are the reason for the next read-only VBA export.

| Worksheet | Tracked write path / destination | Protection implication |
|---|---|---|
| `DAILY_INPUT` | User entry: therapist rows `A7:F16`; patient absence/cancellation rows `A21:E40` for current BF's spec (10 therapist rows, 20 patient rows). Date `B2` and headings are generated metadata. Daily suspension/replacement macros also append/clear operational patient facts in this sheet. | Candidate policy: unlock only the explicit user entry cells; protect date, headings, layout and all other cells. Macros need trusted write access to append/clear sanctioned rows. Reconfirm actual spec from BF and all operational paths before locking. |
| `THERAPIST_DAILY` | Daily render plan emits only cells on this sheet (`render_writeback.py` rejects other sheets); daily action VBA applies validated cell patches; outpatient in-session writer adds/removes patient lines; daily restore/ledger changes rendered lines; workload strip reads it. | Operational projection, not a free-form grid. Protect all cells; trusted macros/backends must write through an explicit controlled path. Buttons must remain clickable, so drawing-object protection settings need deliberate handling. |
| `MASTER_SCHEDULE` | Patient registration bridge rebuilds `A:P` projection rows and formulas; patient program renderer writes treatment display columns D:K; master projection writer writes/clears A:P rows; master toolbar changes view/scroll/row visibility and layout. | Protect manual editing; VBA projection updates need UI-only protection or an explicit trusted writer. Preserve toolbar shapes. |
| `PATIENTS` | Registration bridge and Python registration writer append rows; edit form invokes `UpdatePatientInWorkbook` (the actual implementation is not present in the reviewed form module and must be located in workbook VBA); registration bridge can add metadata headers. | Registry is authoritative. Protect direct grid edits; registration/edit workflows require a controlled writer and read-back tests. |
| `PATIENT_PLANNER` | Identity/status columns A:E are formula projections from PATIENTS. Patient program form writes per-patient therapy time/day/provider fields by resolving header names; row is found by PatientID; atomic snapshot/rollback is implemented. | Protect direct edits; permit the program form only. Ensure its writes and rollback continue under protection. |
| `OUTPATIENT_SCHEDULE` | Python outpatient upsert writes A:F rows; in-session VBA writes A:F; sheet may be created on demand. | Protect direct edits; form/backend is the only sanctioned writer. Workbook-structure lock must account for on-demand creation. |
| `SETTINGS` | Therapist registration writer appends to column A (THERAPISTS_FTH). Patient/outpatient forms read rooms (H), patient statuses (F), time slots (B), days (D), treatment types (E). | Mostly configuration. Column A has a sanctioned registration writer. Other editable configuration ranges are not yet confirmed; do not blanket-unlock or assume no admin editing. |
| `_PY_LISTS` | Daily-input builder regenerates hidden dropdown source lists in A:E and creates workbook defined names. | System-generated, veryHidden. Protect; builder requires a trusted write/refresh path. |
| `DAILY_CHANGE_LEDGER` | In-place VBA appends change rows A:L, updates status/restored timestamp; install path ensures sheet. | Technical ledger, veryHidden; protect and preserve macro write/read-back/restore. |
| `FINALIZED_DAY_SNAPSHOTS`, `THERAPIST_DAY_HISTORY`, `PATIENT_SESSION_HISTORY`, `DAY_HISTORY_EVENTS` | Close-day and amendment VBA unprotects, writes rows, verifies read-back, then veryHides and protects. Amendment also supersedes previous snapshot and has rollback. | Already protected. Preserve exact write/rollback/finalize sequence; any new protection mechanism must not break passwordless trusted unprotect/reprotect semantics. |
| `PATIENT_HISTORY_REPORT` | Patient report renderer unprotects, clears/rebuilds report area, adds buttons, then protects. | Already protected. Buttons currently rely on drawing-object protection being False. |
| `THERAPIST_HISTORY_REPORT` | Therapist report renderer unprotects, clears/rebuilds A:M, adds buttons, then protects. | Was unprotected in initial BF audit. Ensure initial protection on install/open and preserve renderer. Buttons rely on DrawingObjects=False. |
| `DAILY_REPLACEMENT_FACTS` | Replacement workflow creates this sheet if missing; writes A:J, sets veryHidden and protects; restore/rollback clears rows. Not present in the 23-sheet BF inventory at time of audit, so creation is dynamic. | Must be included in dynamic sheet-creation and protection lifecycle. |
| `STUDENTS` | Separate student registration backend writes A:G, but no STUDENTS sheet was present in BF's 23-sheet inventory. | Not part of this BF workbook baseline; include only if student registry is integrated in the release workbook. |

## Sheets with unresolved ownership / workbook-only macros

The following BF sheets have not yet been fully traced to their writers from tracked source:
- `THERAPIST_ATTENDANCE` (RefreshAll calls `RefreshTherapistAttendanceAndConflicts`; implementation is not among the reviewed Python VBA source modules)
- `REPLACEMENTS`
- `REPLACEMENT_LOG`
- `DAILY_LOG`
- `STATISTICS`
- `SESSIONS`
- `CONFLICT_LOG`
- `NEW_PATIENT`

Do not decide unlock ranges for these until the full embedded VBA project is inventoried and any user-facing data-entry contract is identified. Some may be legacy or generated, but that is not proven.

## Protection architecture risks

1. Worksheet `UserInterfaceOnly:=True` is session-scoped in Excel and must be re-applied on Workbook_Open / controlled session refresh.
2. Several tracked VBA paths call `Unprotect` and later `Protect` without a password. A protection retrofit must preserve this trusted macro pattern or migrate it to a common safe wrapper.
3. Some Python COM backends open a fresh Excel instance and write directly to PATIENTS, OUTPATIENT_SCHEDULE, SETTINGS, MASTER_SCHEDULE or STUDENTS. `UserInterfaceOnly` in the user's open Excel instance does not authorize a separate COM instance. These writers need an explicit controlled unprotect/reprotect context (or another verified design).
4. Workbook structure protection is currently False. Several flows create sheets on demand (`FINALIZED_*`, `OUTPATIENT_SCHEDULE`, `DAILY_REPLACEMENT_FACTS`, `_PY_LISTS`, `DAILY_CHANGE_LEDGER`) and hide them. Do not turn on ProtectStructure until dynamic create/rename/visibility paths are handled and regression-tested.
5. DrawingObjects must remain unprotected or be individually managed on user-facing sheets where toolbar and report buttons need to work.
6. Existing technical history sheets are protected with default protection options. Keep them veryHidden and preserve exact rollback semantics.

## Next diagnostic: export workbook-embedded VBA write surfaces (read-only)

Added:
- `Python/tools/audit_vba_write_surfaces.py` commits `f21deed`, `b913472`, `a68c7d8`, `728c50f`, `d1214d5`.
- `Tests/test_audit_vba_write_surfaces.py` commits `bf54332`, `41e365a`.

The script opens BF ReadOnly=True, reads every VBA component/module line, and reports procedure name, module line number, sheet references, writes, protection, formatting, visibility, and workbook-state changes. It never edits or saves the workbook and closes with SaveChanges=False. It has output caps; any "INCOMPLETE" result must be followed by targeted extraction, not treated as a complete audit.

After pull, run:
```powershell
python -m pytest Tests/test_workbook_protection_audit.py Tests/test_audit_vba_write_surfaces.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Protection diagnostics tests failed" }

python .\Python\tools\audit_vba_write_surfaces.py `
  ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H7_20261104_BF.xlsm"
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: VBA write-surface audit incomplete" }
```
Expected: 8 passed (5 protection audit + 3 VBA inventory). Share the complete VBA audit output.

## No protection implementation yet

No worksheet cells have been unlocked/locked and no new Protect/Unprotect calls have been installed. The next step is to reconcile the full workbook-VBA inventory with the tracked source map, decide user input ranges, then create focused tests and a new uniquely named disposable workbook candidate. BF remains frozen evidence; do not patch it or production.


## Exact operator audit received: BF inventory — 2026-11-04

The complete operator file `H7_PROTECTION_AUDIT.txt` was received and reviewed. It reports:
- `ReadOnly=True`
- `ProtectStructure=False`, `ProtectWindows=False`
- 23 worksheets
- `Workbook unchanged: True`
- Audit completed without truncation or COM error.

### Exact protection state

18 worksheets have `ProtectContents=False`:
| Sheet | Visible enum | UsedRange | UsedCells |
|---|---:|---|---:|
| `_PY_LISTS` | 2 (veryHidden) | A1:E93 | 465 |
| `DAILY_INPUT` | -1 | A1:F40 | 240 |
| `PATIENT_PLANNER` | -1 | A1:W550 | 12,650 |
| `DAILY_CHANGE_LEDGER` | 2 (veryHidden) | A1:L3 | 36 |
| `OUTPATIENT_SCHEDULE` | -1 | A1:F1 | 6 |
| `MASTER_SCHEDULE` | -1 | A1:M553 | 7,189 |
| `THERAPIST_ATTENDANCE` | -1 | A1:F18 | 108 |
| `REPLACEMENTS` | -1 | A1:A25 | 25 |
| `REPLACEMENT_LOG` | -1 | A1:H3 | 24 |
| `DAILY_LOG` | -1 | A1:G1 | 7 |
| `STATISTICS` | -1 | A1:E37 | 185 |
| `PATIENTS` | -1 | A1:E550 | 2,750 |
| `NEW_PATIENT` | -1 | A1:D6 | 24 |
| `SETTINGS` | -1 | A1:H32 | 256 |
| `SESSIONS` | 0 (hidden) | A1:H180 | 1,440 |
| `CONFLICT_LOG` | -1 | A1:A2 | 2 |
| `THERAPIST_HISTORY_REPORT` | -1 | A1 | 1 |
| `THERAPIST_DAILY` | -1 | A1:J25 | 250 |

Five worksheets have `ProtectContents=True`:
- `PATIENT_HISTORY_REPORT` — visible; `ProtectDrawingObjects=False`, `ProtectScenarios=True`; no unlocked cells in UsedRange.
- `FINALIZED_DAY_SNAPSHOTS`, `THERAPIST_DAY_HISTORY`, `PATIENT_SESSION_HISTORY`, `DAY_HISTORY_EVENTS` — all veryHidden; `ProtectDrawingObjects=True`, `ProtectScenarios=True`; no unlocked cells in UsedRange.

Every sheet's reported `UsedRange.Locked` property is True. For the 18 unprotected sheets this is not effective edit protection because `ProtectContents=False`. Every sheet reports `EnableSelection=xlNoRestrictions`. All displayed formatting/insertion/deletion/filter/sort allowances are False.

### Interpretation and scope correction

- The exact audit confirms the prior 18/5 count and adds each UsedRange, visibility enum, and cell count. It does not identify intended user-entry cells.
- The workbook's structure/windows remain unprotected. Keep this so for now: tracked flows create/hide sheets dynamically.
- `THERAPIST_HISTORY_REPORT` is visible but currently blank (UsedRange A1) and unprotected. The report renderer protects it only after a successful render; the installer also left the blank sheet unprotected. This is a genuine gap because users can type into the blank report sheet before opening the report.
- The tracked therapist report renderer also had an error-path gap: after `ws.Unprotect`, an exception before its normal final Protect could leave the report sheet unprotected. This is now fixed in code as described below.
- `SESSIONS` is hidden (Visible=0), not veryHidden. Do not silently change its visibility in this protection slice.
- Do not assume the other 17 unprotected sheets are all data-entry surfaces. The intended human edit zones remain unresolved except DAILY_INPUT.

## Targeted therapist-report protection fix

Test-first changes:
- `45ed8c9` adds regression assertions that the therapist report is protected after install and in the VBA error cleanup path.
- `9a1574e` updates `THERAPIST_HISTORY_MODULE_CODE` so its `CleanUp` label always re-applies `Protect(DrawingObjects:=False, Contents:=True, Scenarios:=True)`, including errors after unprotect.
- The same `9a1574e` update makes `install_therapist_history_report` protect the blank report sheet immediately after creating its buttons and verifies `ProtectContents=True` and `ProtectDrawingObjects=False`.

This is deliberately limited to `THERAPIST_HISTORY_REPORT`; it does not implement broad protection on the other 17 unprotected sheets. The user-facing report buttons remain clickable because DrawingObjects protection stays False. Native smoke verification is still required on a fresh workbook after focused tests pass.

## VBA audit count interpretation

The earlier VBA scanner reported 1,146 matching source lines across 64 components. Treat 1,146 as **lexical candidate lines**, not 1,146 proven writes: the scanner intentionally also reports sheet references, formatting properties, reads, visibility and protection calls. A line-level semantic classification and procedure-to-sheet mapping is still required. The scanner output alone is not authority to choose unlock ranges.

## Next exact gate

After pull:
```powershell
python -m pytest Tests/test_therapist_history_ui.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Therapist report protection tests failed" }
```

If green, prepare a new uniquely named H7 protection smoke from the trusted build path (never reuse BF), install H7 UI, run read-only audit and manually verify:
1. blank therapist report sheet is protected immediately after install;
2. therapist report button still opens and refreshes daily/weekly/monthly/overall;
3. report buttons remain clickable;
4. force a controlled report-render failure after Unprotect in a disposable test and verify the sheet is re-protected;
5. patient report behavior and previous-screen navigation remain unchanged.

Do not begin locking other sheets until all 18 unprotected sheet ownership paths are classified and intended manual edit ranges are approved by evidence.
