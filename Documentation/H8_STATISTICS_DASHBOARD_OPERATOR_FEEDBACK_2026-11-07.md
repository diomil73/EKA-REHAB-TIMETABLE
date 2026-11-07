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
