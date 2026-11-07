# H8 — Statistics dashboards, clarity and operational KPIs
Date: 2026-11-07
Branch: work/history-statistics-rebuild
Status: ACCEPTANCE FINDINGS CONFIRMED; specification and source investigation checkpoint. No source implementation or native H8 smoke has yet been accepted.

## Operator's BI findings (authoritative requirements)

1. RefreshAll is currently intrusive: individual modal OK dialogs appear for each refreshed item. Question whether the button should exist at all. If it stays, it should display **one** final notification only, «Όλα ενημερώθηκαν», and never demand multiple confirmations.
2. Patient reports must show **patient names**, not bare numeric IDs.
3. Therapist statistics currently show the same numbers for daily, weekly, monthly, overall, and «Ημέρες» shows 1. Operator needs explanation of the history coverage and meaningful distinct period views.
4. Therapist report must clearly expose **workload**, **therapist absences**, replacement/coverage facts and all previously agreed operational indicators.
5. Patient **participation percentage** and other actual statistical outcomes must be shown, not merely raw count columns.
6. Clear, readable **charts** are required for overall and per-person (patient / therapist) dashboards, with an at-a-glance overview.

### Acceptance correction

BI H7 source tests (33/33), integrated native install, persisted audits and workbook protection audit are GREEN, but **BI is NOT product/UX accepted**. These issues are operator-observed during manual exploration. The earlier plan to immediately proceed to full worksheet protection is superseded. Stop general protection implementation until the analytics UX and metrics have been addressed and acceptance re-evaluated. Do not overwrite BI; freeze as a traceable diagnostic UX-gap smoke. AZ remains trusted/frozen; BF green functional evidence but not H8 analytics acceptance; BG/BH failed/frozen.

## Verified tracked source findings

- `Python/tools/install_patient_history_ui.py`: installer positions and retains a canonical `ekaToolbar_RefreshAll` with macro `RefreshAll`; the underlying VBA is the existing `Module3.Public Sub RefreshAll()`. Refactor **that macro**, not the shape-only installer, for single-summary feedback. Never suppress or convert failures into a false «Όλα ενημερώθηκαν». Inspect each delegated procedure, screen state, dependency and VBA error handling.
- `Python/tools/export_patient_history_ui.py`: TSV header explicitly exports `PatientId` and `_display_patient_id` applies two-digit padding for display. The report renderer in `Python/rehab_excel/patient_history_report_vba.py` renders that field under «Ασθενής» using only 9 TSV columns. `Python/rehab_excel/patient_registry_source.py` already reads `Patient(patient_id, display_name, ...)` from `PATIENTS`. The name must be joined via stable patient identifier (not derived from cell position); protect against duplicate/unknown IDs, names changed after history, Unicode/Greek, zero-activity rows, and missing historical identity data. Never substitute a guessed name.
- `Python/rehab_core/therapist_history_reports.py`: distinct daily/weekly/monthly/overall aggregators already exist, and daily reports are restricted to CURRENT_FINAL finalized snapshots. They use recorded historical days only and avoid double counting amended snapshots. `Python/tools/export_therapist_history_ui.py` exports all four. If only **one** finalized date is present, reported days=1 and all four views may legitimately have the same totals. The full BI report rendering has not yet been independently reproduced, so do not claim there is no separate UI bug.
- `Python/rehab_excel/therapist_history_report_vba.py` already has columns `Προγρ. φόρτος` and `Πραγμ. φόρτος`, mapped from `planned_capacity_utilization` and `actual_capacity_utilization` as percents, but rendered in a wide 13-column table. Missing capacity results in blank percentage. Data fields also include replacements received, releases and treatment suspensions. These are **not a complete, clearly described therapist absence metric**. `THERAPIST_DAY_HISTORY` schema tracks slot/session counts and suspensions, but no standalone therapist absence count/duration; a new authoritative provenance/definition is needed before displaying one.
- `Python/rehab_excel/patient_history_reports.py` has scheduled, performed, patient_absences, cancellations, replacements per report period; participation KPI is not included yet.
- Existing read-only stats core: `Python/rehab_excel/historical_statistics.py`, `patient_statistics.py`, `period_statistics.py`. Maintain the CURRENT_FINAL and version-aware contracts and never count older amendments as extra treatment days.

## Required product behavior

### A. Refresh interaction

- Determine whether RefreshAll is still needed given report-mode exporters already refresh when buttons are clicked. Keep only if it updates other live operational views beyond report renders, with an explicit accessible purpose.
- One **terminal** success dialog at most: «Όλα ενημερώθηκαν». Intermediate modal success dialogs are forbidden. A failure should produce one clear summary explaining which operation failed; never indicate complete success if something failed.
- No lost Save, CloseDay, Amend, manual controls or dangerous implicit recalculations; VBA session flags restored on errors; avoid reentrant refresh.
- Regression tests for sequential calls, no duplicate modal prompts, failure summary and correct OnAction.

### B. Patient identity and KPIs

- Present full `display_name` as the primary label, optionally retain stable ID as a discreet secondary discriminator for namesakes.
- Resolve ID to current registry strictly; if the patient was discharged or no longer in current registry, use reliable historical-name snapshot when available, otherwise explicitly show «Άγνωστο όνομα (κωδικός …)». **Never infer or invent names**.
- Name changes: current-registry display names are not a historical name guarantee. Historical provenance should be clear.
- Participation candidate formula, subject to cancellation policy verification: `100 × performed / eligible_scheduled`. Count ACTIVE and REPLACED delivered sessions exactly once. Define eligible_scheduled and treatment suspensions/clinic cancellations separately before release. In parallel show performed, scheduled, patient absences, cancellations, replacements, period coverage. Zero denominator: `— (δεν υπάρχουν επιλέξιμες συνεδρίες)` not 0% or 100%.
- Distinguish patient-attributable absence/no-show, clinical cancellation and therapist unavailable when source supports it. Do not conflate patient absence with all cancellations.
- Summarize totals and per-person trends, using finalized current-version data only.

### C. Therapist KPI definitions

- Visually prominent: available/capacity slots, planned occupied slots, delivered slots, utilization %, planned vs actual sessions, replacements given/received where reliably derived, released slots, suspensions and recorded therapist absences.
- Historical therapist-absence count/duration needs a source trace: day absence events, therapist attendance/absence journal or finalized evidence. If absent, mark «Μη διαθέσιμο ιστορικό απουσιών», not zero. Cross-check absence vs suspension semantics; do not call suspensions «απουσίες».
- Workload utilization formula (existing): occupied slots / capacity slots when capacity >0. If capacity missing show `—` with reason. Distinguish percentage utilization from sessions per day; avoid double count of shared slots.
- `Ημέρες` should read **«Καταγεγραμμένες οριστικοποιημένες ημέρες»**, not calendar days. Display coverage: first and last finalized dates, number of finalized days, and a note «Τα στατιστικά βασίζονται μόνο σε οριστικοποιημένες ημέρες». If exactly one day exists, state «Διαθέσιμη 1 οριστικοποιημένη ημέρα, γι’ αυτό οι περίοδοι μπορεί να έχουν ίδιες τιμές».
- Clearly show active period boundaries and compare aggregate totals to the relevant source days. Tests need 1-day and multiple-day/multiple-week/multiple-month cases and v1→v2 amendment.

### D. Charts and clear dashboards

- Both high-level overall dashboard and each individual's patient/therapist profile need a prominent native Excel chart area, large Greek titles, visible legend, unit and time period, no chart overlapping report controls, data table or shell.
- Overview suggested: patient completed vs absent vs cancelled sessions; therapist capacity versus occupied slots and utilization; trend by week/month from CURRENT_FINAL history.
- Patient profile: delivered vs missed/eligible sessions and participation %; weekly/monthly trend where coverage allows.
- Therapist profile: planned vs actual load and utilization %; substitutions/suspensions/absences as separate labeled breakdown where credible.
- Charts should be native Excel objects or safely rendered workbook UI, refreshed/rebuilt idempotently on mode switch, protected against manual edits while **remaining viewable**. Ensure chart shapes do not disable navigation buttons, and no duplicate charts accumulate.
- Denominator, source coverage, legend and date range must be legible at a glance. Missing data must display an explicit empty state, not fabricate a graph. Native tests validate 1 finalized day, multiple periods and amendments.

## Delivery order / gates (no automatic migration to production)

1. **Source audit**: enumerate VBA Module3 RefreshAll body and called macros; trace therapist absence historical source and patient identity provenance; formalize KPI denominator and source coverage.
2. **Backend correctness tests** for patient-name ID joins, duplicate identities, period aggregation, participation edge cases and missing/partial therapist capacity and absence provenance.
3. **Dashboard / Excel VBA UI**: overview, individual selection, charts and captions, report protection and report toolbar integration. Keep historical storage and daily save semantics unchanged.
4. **Source regression gate**: existing 33 H7 tests + focused H8 tests + appropriate full-history tests.
5. **Fresh uniquely named native XLSM** from trusted authorized source after source tests green, with installer persisted reopen and full read-only protection audit. Do not overwrite frozen AZ/BF/BG/BH/BI or any existing destination.
6. **Operator manual acceptance** across daily, weekly, monthly, overall for both populations, chart accuracy and clickability, missing data and update feedback. Only then resume general worksheet protection write-path audit.

No new H8 smoke, acceptance claim, or production workbook edit is authorized by this document. This is a specification / root-cause investigation checkpoint, not a code fix.

## Implementation checkpoint — first H8 code slice committed (2026-11-07)

The operator explicitly requested immediate implementation. These **source-only** commits exist on `work/history-statistics-rebuild`:
- `e355075`: patient TSV exporter obtains stable current-registry `display_name`, refuses duplicate normalized registry identifiers, preserves unknown-name fallback for absent historical IDs, and adds `CompletionPct`.
- `bfa257b`: patient VBA report now renders names and `Συμμετοχή %` in a 10-column table; it does not display patient codes as the primary label.
- `70f0c30`: H7 patient VBA schema test adjusted from 9 to 11 TSV fields.
- `8578ce1`: new H8 unit tests for identity matching, collision stop, completion edge cases, unknown names, workbook-byte immutability.
- `38c6f3d`: therapist report explicitly labels recorded finalized days and notes single-day period equality.
- `b894097`: first native patient overview chart (completed/absent/cancelled), idempotent one-chart pattern, above 10-column table in protected report.
- `e8faa2f`: patient chart source contract test.
- `38cc657`: first native therapist load overview chart (planned/actual occupied slots), idempotent and above history table.
- `b929cc9`: therapist layout regression tests adjusted, first chart test.
- `7ac784e`: existing patient mode-switch tests updated for 10 columns.
- `a72a448`: existing H7 patient exporter tests updated for name join and completion metric.

**Status:** source changes staged/committed, **no operator source-test result yet**, **no native H8 XLSM yet**. Never apply to BI or alter frozen proof. Chart work here is overview-only, not complete per-person drilldown; abstain from claiming the full dashboard accepted. Genuine therapist absence metrics and per-person charts still require source evidence and implementation.

### Accuracy note for first KPI

`CompletionPct = 100 × performed / scheduled`, with zero scheduled -> empty/undefined, counting cancellations within scheduled. It is explicitly a *completed versus originally scheduled* metric, not an adjusted-eligible attendance KPI. The final product may need a second clinically agreed patient participation metric after cancellation classifications are verified. Never label missing denominator as 0%.

### Known existing RefreshAll implementation

`Documentation/CHECKPOINT_H7_REFRESHALL_RETAINED_AND_TOOLBAR_MIGRATION_2026-11-04.md` contains the read-only VBA procedure from a frozen workbook:
`Module3.RefreshAll` calls six operational routines: `RefreshTherapistDaily`, `RefreshTherapistAttendanceAndConflicts`, `BuildReplacementSuggestions_V4`, `ColorInfectiousInMaster`, `ColorReplacementMarks`, `RestoreReturnedTherapists`, then gives a terminal `MsgBox "Όλα ενημερώθηκαν."`. It is therefore not redundant and must not be removed. Repeated intervening dialogs most likely arise in delegated routines; the exact call chain needs native VBA source inspection before suppressing intermediate prompts. Do not blindly remove dialogs from operational macros that also run independently, or misreport success on errors.

### Next safety gate (Excel fully closed)

```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }

python -m pytest `
  Tests/test_h8_patient_names_and_participation.py `
  Tests/test_patient_history_ui_h7.py `
  Tests/test_patient_history_reports_h7.py `
  Tests/test_h7_patient_mode_switch_and_toolbar_fix.py `
  Tests/test_therapist_history_ui.py `
  Tests/test_h7_integrated_toolbar_and_patient_order.py `
  Tests/test_h7_history_previous_screen_navigation.py -q

if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: H8/H7 source regression tests failed" }
```

This gate is expected to exercise 6 new H8 tests plus existing H7 suites. Do **not** create a native workbook if any test fails. Send complete failing traceback; correct source and retry source gate. If green, then verify additional relevant history tests before a new H8 smoke, and run full read-only audit before any native manual click.

**Open work:** exact RefreshAll suppression strategy; absence provenance; full patient/therapist individual chart and KPI drilldowns; clear multi-period data-coverage display; native COM chart render proof. This first code slice is not a release-ready implementation.

## Operator proof — 2026-11-07, H8 per-person chart source gate GREEN

The operator confirmed the **seven-file targeted regression suite completed: `52 passed in 5.51s`**. This result is later than the per-person chart source commits `57d7220`, `2dca8ac`, `fbde9a1`, `a3ecbcc`, `186ce00`, and `5531c28`. Thus the current committed patient/therapist native chart source contracts and H7 report regressions passed source tests. It does **not** prove chart behavior in live Excel, VBA procedure interactions, read-only protection after chart interaction, single-message RefreshAll behavior or genuine therapist-absence KPIs. No H8 native workbook has been created or accepted. Preserve AZ/BF/BG/BH/BI frozen. Priority is remaining source audit / targeted remediation and then a fresh uniquely named native workbook with complete persisted state audit before manual tests.

## H8 RefreshAll diagnostic bridge staged after 52 GREEN (2026-11-07)

The operator's previous seven-file gate returned **52 passed in 5.51s**. Since then, source commits `00eb3d4` and `012a936` added `Python/tools/audit_h8_refresh_dialogs.py` plus three focused tests. This is a new **read-only** VBA dependency/dialog candidate audit. It does not execute macros, change Excel, patch Module3 or claim that lexical MsgBox sites necessarily appeared at runtime. It recursively inspects `RefreshAll` and its six known operational delegates; missing/ambiguous VBA definitions fail closed. Every exit closes the workbook `SaveChanges=False`.

**Next exact gate with Excel completely closed** (new tests were not part of the operator's 52-pass confirmation):

```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }

python -m pytest Tests/test_h8_refresh_dialog_audit.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: H8 refresh audit source tests failed" }

$proof = ".\\Excel\\previews\\UNIFIED_PATIENT_HISTORY_H7_20261104_BI.xlsm"
if (!(Test-Path -LiteralPath $proof)) { throw "SAFETY STOP: frozen BI not found" }

python .\\Python\\tools\\audit_h8_refresh_dialogs.py $proof
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: H8 RefreshAll audit incomplete or failed" }
```

Return the full diagnostic output, particularly each `DIALOG CANDIDATE` and `MISSING PROCEDURE` line. This is non-mutating inspection of **frozen BI**; do not click any refresh buttons, save BI, run installers or create the next H8 workbook. Once authoritative call sites are established, make targeted source-safe dialog cleanup while preserving all six operational effects and failure notifications. Analyst must not blindly suppress intermediate error or confirmation prompts.

**Other pending H8 work:** therapist absence history provenance/definitions, full individual trend charts across multiple finalized periods, new source and Excel COM native chart proof, patient metric cancellation-policy semantics. Current source status is 52-pass gate for previous per-person chart slice; **the new refresh diagnostic source tests are still pending**.

## Operator's authoritative native RefreshAll dialog audit (2026-11-07)

The operator ran `Tests/test_h8_refresh_dialog_audit.py`: **3 passed in 0.68s**. Then the read-only tool inspected frozen BI and reported `ReadOnly: True`, `Workbook unchanged: True`, `H8 REFRESH DIALOG AUDIT COMPLETE`. The exact candidate success dialogs:
- Module3:17 `RefreshAll`: `MsgBox "Όλα ενημερώθηκαν.", vbInformation` (the only desired terminal success).
- Module1:380 `RefreshTherapistDaily`: `MsgBox "Το THERAPIST_DAILY ανανεώθηκε.", vbInformation`.
- Module2:229 `RefreshTherapistAttendanceAndConflicts`: `MsgBox "THERAPIST_ATTENDANCE και CONFLICT_LOG ανανεώθηκαν.", vbInformation`.
- Module4:217 `BuildReplacementSuggestions_V4`: `MsgBox "Οι προτάσεις αναπλήρωσης v4 δημιουργήθηκαν.", vbInformation`.
- Module4:336 nested `AddReplacementButtonsV4`: `MsgBox "Κουμπιά που δημιουργήθηκαν: " & createdCount, vbInformation`.
- Module10:68 `RestoreReturnedTherapists`: `MsgBox "Ολοκληρώθηκε ο έλεγχος επαναφοράς.", vbInformation`.

All six expected operational refresh procedures and nested call candidates were found; these are **candidate sites**, not proof each one triggers in every execution. The remaining non-message delegates `ColorInfectiousInMaster` and `ColorReplacementMarks` were found. Frozen BI was not edited.

## Source-only RefreshAll remediation checkpoint

New commits:
- `8b1f148` `Python/tools/install_h8_single_message_refresh.py`: pure preflight exact contract validation of original six-call `Module3.RefreshAll`; adds a public scoped flag `EKA_SilentRefreshSuccess`, one terminal success, one error handler with Excel session state restoration; wraps the five exact delegate-success dialogs only while the total refresh is underway, preserving independent success messages and other `vbCritical` messages. Fails closed if VBA procedure text or exact dialog occurrences differ. Installs only into a fresh disposable destination, followed by read-only persisted VBA source equality verification.
- `b4828fd` `Tests/test_h8_single_message_refresh.py`: source contracts and fail-on-drift, duplicate messages and reinjection.
- `46cbd0e`, `6debc43` `Python/tools/prepare_h8_dashboard_smoke.py`: one-shot fresh XLSM builder using trusted AZ via the existing H7 installer chain and then the H8 refresh installer. Does not patch frozen BI.
- `c496c82` `Tests/test_h8_dashboard_smoke_builder.py`: sequencing / refuse-overwrite / freeze-on-failure tests.

**Important:** These code changes have NOT passed a new operator test gate. None has been run in native Excel yet, and the H8 workbook has NOT been created. The existing 52-pass gate occurred before these commits. Do not claim verified single-dialog UX or broader dashboard acceptance before a new source test and native smoke. No change to historical store, clinic data, or frozen files was made.

### Next source gate

With Excel fully closed, from repo root:

```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }

python -m pytest `
  Tests/test_h8_single_message_refresh.py `
  Tests/test_h8_dashboard_smoke_builder.py `
  Tests/test_h8_refresh_dialog_audit.py `
  Tests/test_h8_patient_names_and_participation.py `
  Tests/test_patient_history_ui_h7.py `
  Tests/test_patient_history_reports_h7.py `
  Tests/test_h7_patient_mode_switch_and_toolbar_fix.py `
  Tests/test_therapist_history_ui.py `
  Tests/test_h7_integrated_toolbar_and_patient_order.py `
  Tests/test_h7_history_previous_screen_navigation.py -q

if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: H8 source regression tests failed" }
```

Only after this gate is green may a **fresh one-shot** candidate be created as `Excel/previews/UNIFIED_PATIENT_HISTORY_H8_20261107_BJ.xlsm` from frozen trusted `Excel/previews/UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm`, provided `BJ` does not already exist. The new builder is `python Python/tools/prepare_h8_dashboard_smoke.py --source ... --output ...`. After new build, repeat full independent `audit_workbook_protection.py`, `inspect_h7_integrated_ui.py`, and new `audit_h8_refresh_dialogs.py` (the latter lists lexical dialog candidates still present, intentionally under scoped flags). Freeze BJ on any failure; do not retry/patch an existing BJ. No manual interaction until audit passes. During native manual smoke, test RefreshAll **once** (six functions must execute), both reports across all four modes, names, KPI, aggregated and selected-person chart, mode switching, button placement, protection and return.

**Open H8 acceptance gaps**: final therapist absence provenance KPI and full longitudinal individual trend visualization remain to be built. First H8 chart release is still a trial, not full product acceptance. Full worksheet protection of the other 17 sheets remains unauthorized.

## Operator source gate GREEN — 2026-11-07; BJ native proof authorized

The operator confirmed the exact combined H8/H7 test command after the guarded RefreshAll installer and one-shot builder commits: **63 passed in 5.63s**. This supersedes the previous pending-test status for that source gate.

**Disposition:** H8 source regression gate GREEN. Authorize one fresh native candidate `Excel/previews/UNIFIED_PATIENT_HISTORY_H8_20261107_BJ.xlsm`, created only once from trusted frozen AZ using `Python/tools/prepare_h8_dashboard_smoke.py`. BI and earlier evidence remain untouched. This is **not** native VBA or chart success and not final H8 acceptance.

Exact one-shot PowerShell command (close Excel first):

```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }

$source = ".\\Excel\\previews\\UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm"
$output = ".\\Excel\\previews\\UNIFIED_PATIENT_HISTORY_H8_20261107_BJ.xlsm"
if (!(Test-Path -LiteralPath $source)) { throw "SAFETY STOP: AZ source missing" }
if (Test-Path -LiteralPath $output) { throw "SAFETY STOP: BJ already exists; do not overwrite or reuse" }

python .\\Python\\tools\\prepare_h8_dashboard_smoke.py --source $source --output $output
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BJ build failed; freeze BJ and do not retry" }

python .\\Python\\tools\\audit_workbook_protection.py $output
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BJ protection audit failed; freeze BJ" }

python .\\Python\\tools\\inspect_h7_integrated_ui.py $output
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BJ integrated toolbar audit failed; freeze BJ" }
```

Expected native build: H7 copy and installation sequence verified; `H8 REFRESH ONE-MESSAGE INSTALL OK`, `Persisted VBA reopen verification: OK`, `H8 DASHBOARD PREPARATION OK`. Protection audit must report both patient and therapist report `ProtectContents=True`, `ProtectDrawingObjects=False`; all four technical history sheets protected/veryHidden; `ReadOnly=True` and `Workbook unchanged: True`. Toolbar audit should confirm canonical RefreshAll and two statistics buttons, no overlaps or missing macros. No manual button clicks or Save/Exit until the complete read-only output is reviewed.

If any step fails, do not re-run against the same BJ; report full terminal output. If BJ existed before this command, stop at the safety gate and preserve it. After audit GREEN, manually smoke test native charts, report modes and one terminal RefreshAll success only (failure messages retained). Charts and success-dialog behavior have **not** yet passed native proof.

Open H8 scope remains genuine therapist absence provenance, clinical denominator policy, and longitudinal individual trends, as well as broad worksheet protection. Do not conflate source green with delivery acceptance.
