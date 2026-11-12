# Diomil functional demo: outpatient end, inpatient permanent discharge, therapist/student roster lifecycle

**Operator instruction (2026-11-12):** Pilot is synthetic-data-only; actual 2027 program NOT imported. Operator enters outpatient appointments through the existing outpatient management menu. Cosmetic finish is postponed. All patient, roster and daily logic must be proven in one functional demo before clinical use.

## Required behavior

1. **Outpatient program ends:** Make an explicit end-program action reachable from the existing outpatient schedules menu. Effective date is recorded; stop generating recurring outpatient bookings after that date (inclusive end date means last allowed treatment date). The patient must no longer occupy future daily timeslots after the stop. Preserve prior/closed finalized sessions, original program and audit trail. Read-only preview of affected upcoming bookings before committing. Do **not** equate program end with an absence/cancellation for an individual day.
2. **Inpatient permanently discharged:** Make an explicit permanent discharge workflow via the patient management UI. Record effective discharge date; remove patient from active future inpatient roster/timeslot generation from the next applicable day, without deleting patient identity or changing historical finalized visits, absent reasons or records. Protect against reuse of formerly occupied appointments. Existing statuses/workflows must be reused where applicable; do not introduce new patient absence types. Return/re-admission, if needed, must be an explicit separately validated action, not silent reactivation.
3. **Add and remove therapists:** Use one managed therapist roster, with add, edit and **deactivate/retire (not hard-delete)** actions. Newly active therapist must appear in all relevant ΦΘ/robotics selectors and planning capacity. Retirement must block if the therapist still owns future appointments unless they are explicitly transferred or resolved. Preserve stable historical therapist IDs, names and archived work.
4. **Add and remove students:** Use the separate `STUDENTS` registry. Student onboarding requires StudentID, name, start/end placement dates and supervisor where applicable. Student end/withdraw/deactivation must remove active choices after the effective date, identify affected recurring assignments, and drive safe takeover/reassignment. Preserve identity/history. Existing `student_transition` is a **read-only suggestions planner**, not proof of a committed end-placement workflow.
5. **Daily consistency:** Active census, capacity, `MASTER_SCHEDULE`, `THERAPIST_DAILY`, outpatient menus and statistics must reflect the same effective schedule. One shared slot counts once for capacity. No finalized historical rewriting from a roster change. All writes only through validated native Excel/COM transaction with rollback/reopen/read-back as applicable.

## Existing code inventory inspected

- Outpatient schedule UI: `Python/rehab_excel/outpatient_schedule_manager_vba.py` presents listed schedules, create and edit actions. **No end-program action found** in the inspected manager source.
- Student registry: `Python/rehab_excel/student_registry.py`, `student_registry_schema.py`, `student_registration_form_vba.py`. `Python/rehab_core/student_transition.py` provides read-only end/hand-off suggestions but **does not itself write a permanent transition**.
- Patient state: `Python/rehab_core/models.py` has existing patient types and statuses, which must be consulted rather than multiplied.
- Staff availability/therapist registry: code exists across scheduling and settings; a verified manager for safe add/retire has **not been established** by this review. Do not claim removal already complete.
- Operator enters real outpatients through existing menu **after go-live**, not by importing the external-patient sheet of the uploaded workbook.

## Acceptance tests to implement before demo claim

- Ending an outpatient program leaves past history, eliminates future appointments, frees impacted slot and prevents reappearance on refresh/reopen.
- Discharging an inpatient excludes future recurrence and current active listings after effective date; finalized past events unchanged; no attendance-as-absence fabricated.
- New therapist/student appears in appropriate choice lists; retiring assigned provider is blocked or requires explicit conflict/reassignment resolution; historical IDs remain stable.
- Student placement-end transition handles all affected recurring bookings and leaves no orphaned active assignments.
- A complete synthetic sequence: register -> schedule -> conflict -> absence/replacement -> return -> end/discharge -> save/reopen -> history/statistics. Confirm operational screens agree.
- Must pass Python tests, isolated native Excel smoke, and independent read-only audit. No modification of frozen CC or real workbook and no 2027 patient roster import while gate is closed.

**Status:** requirements frozen; implementation and native tests pending. This documentation is not evidence of functionality.


## 2026-11-12 — Latest operator priority: correct real-data demo MASTER first, then lifecycle mechanisms
The operator explicitly changed the work order: **first transfer today's real patient/therapist/schedule data accurately into a NEW, isolated demo MASTER; only after that finish the remaining lifecycle mechanisms.** Earlier synthetic-only demo direction is superseded. Source remains ONLY newest clear-column worksheet `ΠΡΟΓΡΑΜΜΑ`; excluded `Εξωτερικοί Ασθενείς` will be manually entered via outpatient menu. No actual real-data .xlsm has yet been modified.

Read-only actual source analysis generated a local, sensitive review workbook `DIOMIL_ΕΛΕΓΜΕΝΗ_ΜΕΤΑΦΟΡΑ_ΠΗΓΗΣ_ΠΡΟΣ_MASTER_20261112.xlsx` in the user conversation only, **never commit to GitHub**. It contains 106 candidate source patient rows (54 A, 52 B), **764 candidate weekday/time specialty rows**, **61 source cells held for missing/ambiguous details**, and **11 possible same-patient same-day same-time cross-specialty overlaps** (some correspond to multiple held per-day records). These are **not accepted appointments**. Counts are based on the latest upload, supersede earlier bad parser outputs; they are not independently confirmed census or provider capacity. 29 of the 61 have psychology provider/text but no clock and cannot reserve a timeslot. Pool column V is empty. One robotic record has likely interchanged weekday/provider values but is held, not repaired by guessing. Do not silently import conflicts or use old candidate count as truth.

New tested-by-construction source contract `Python/rehab_core/programme_mapping.py` at `3ac001d`, regression suite `Tests/test_clear_programma_mapping.py` at `b552f19`. It recognizes Excel fractional times such as `0.3854166667 = 09:15`, distinguishes numeric `11.3 = 11:30`, handles ordered slash combinations `11.30/12.15 | ΔΕ/ΠΕ | ΣΑΚΚΟΥ/ΓΑΒΡΑΣ`, requires therapist only for ΦΘ/Ρομποτική and blocks misaligned groups. **Windows pytest and native Excel proof not yet executed; do not declare these tests green.**

The current `PATIENT_PLANNER` stores one treatment-time/weekday combination per specialty and cannot silently absorb different time/provider on different weekdays in that same cell. **Required next architecture**: a lossless dated/weekday-slot backing source for the daily builder with explicit patient ID, specialty, weekday, clock, provider when required, source-cell provenance; keep MASTER visual projection unchanged. Native Excel/COM import into a **fresh** .xlsm only after safe roster identity mapping, cross-specialty conflicts, capacity and history isolation gates; never modify frozen CC or source. All patient identities/clinical data remain out of repo and nonapproved cloud. No real-data deployment claim until native read-back and operator trial.


## 2026-11-12 — Granular weekday storage is integrated at source/read/audit level
To unblock accurate day-specific source mapping, implemented optional **`PROGRAM_WEEKLY_SLOTS`** backing sheet. It is an **additional technical source**, not a new visual MASTER layout. Code `Python/rehab_excel/weekly_slots_source.py` (`0c2be65`, strengthened `84a6e51`) requires exactly `PatientID,Treatment,Weekday,StartTime,Therapist,SourceCell`; checks exact currently known patient IDs, legal specialty/weekday, valid clock, source-cell provenance, mandatory therapist for ΦΘ/Ρομποτικό, and same-patient same-day same-clock collisions. Other specialties can retain an empty therapist. Strict failure on inconsistent rows. **No direct real-data writing to this sheet has occurred.**
`Python/rehab_excel/reader.py` now reads the optional source and replaces the corresponding **patient+treatment** legacy PATIENT_PLANNER single-cell entry with individual weekday entries, preserving the others. Downstream `materialize_sessions_for_date` already operates over `BaseScheduleEntry`; no new MASTER visual columns required. Its source audit also evaluates the effective entry counts and detects invalid granular source rather than double counting legacy and granular rows (`816eb1f`).
Source-specific regressions in `Tests/test_weekly_slots_source.py` (`72cd6de`, `c546d5e`): missing optional sheet backward compatibility, conditional override, date-specific therapist assignment, provider-optional specialty, bad identity/day/provider/source and collision fails closed, effective audit counts.
**All are source-only commits. Tests in authoritative Windows working copy and native Excel/COM write+reopen are NOT yet run.** Frozen H10 CC and current patient workbook unchanged. Do not label the current real-data candidate import as finished. Next: green targeted tests, prepare a verified source-row to explicit patient ID alias map with no guesses, native Excel one-shot fresh-copy ingest; block any unresolved patient/provider identity or patient conflicts before opening candidate for clinical trial.


## 2026-11-12 — Implemented import admission safety gate (not a workbook import)
- `Python/rehab_excel/trial_admission_gate.py` and `Tests/test_trial_admission_gate.py` implement a read-only fail-closed admission verdict over the **local-only** `DIOMIL_ΕΛΕΓΜΕΝΗ_ΜΕΤΑΦΟΡΑ_ΠΗΓΗΣ_ΠΡΟΣ_MASTER_20261112.xlsx` staging review. Inputs are exact source clinic/row/name, explicit confirmed PatientID, proposed weekday placements, unresolved source cells and potential overlaps. Sensitive identity material stays local; the verdict contains no patient names and suppresses released entries when blockers exist.
- `Python/tools/check_trial_admission.py` is a privacy-conscious read-only local CLI. It reports counts and `SAFETY STOP` without writing any workbook or clinical data to GitHub.
- **Checked actual staging workbook sheet names and first-row headers** against the implemented contract. Currently the source's `Αντιστοίχιση PatientID` is blank and `Επιβεβαίωση` is `ΕΚΚΡΕΜΕΙ`; known unresolved cells and overlap rows also remain. Thus **import MUST STOP** until independently reconciled. There is no clinical workbook or finalized master with the newly supplied real patient roster yet.
- Previous 41 targeted pytest cases passed on operator's Windows machine before this latest gate was added. The new admission-gate tests have **not yet been run there**. Do not imply otherwise.
- Next development: controlled identity reconciliation workflow, resolve or quarantine ambiguous source cells without silently dropping material, then secure native-Windows Excel write to a fresh isolated demonstration workbook with full read-back, no CC modification.


## 2026-11-12 — Provider shared-slot preflight on actual staging, safety gate updated
Implemented `trial_admission_gate.py` admission blocker for simultaneous **same-provider, same-weekday, same-time, different-patient** entries, plus positive/negative synthetic tests (`ddd917d`, `368612e`). Read-only independently inspected the private staging review, finding **77 candidate shared provider/day/time keys** involving multiple different patients (79 additional patient placements). Breakdown by treatment label sets: **54 ΦΘ-only, 13 Εργο-only, 7 ΦΘ+Ρομποτική, 2 Ρομποτική-only, 1 Ψυχολογία-only**. These are NOT automatically mistakes: the clinical workflow may permit **shared/group appointments**, which are counted once as an occupied therapist timeslot. Thus the currently conservative blocker is correctly fail-closed but **not a final group-session validator**. Do not auto-delete, move or reject these source entries, and do not mark the real staging as import-ready. Build or connect an **explicit, audited shared-slot approval mechanism** that distinguishes intended multi-patient simultaneous sessions from accidental double-booking, enforces provider slot-capacity by distinct timeslot rather than raw patient count, and keeps patient attendance records distinct. The 77 collisions are separate from the 11 same-patient cross-specialty overlaps; don't conflate them.


## 2026-11-12 — Explicit shared-provider slot approval gate implemented
The admission validator now supports optional **local-only** worksheet `ΕΓΚΡΙΣΕΙΣ_ΚΟΙΝΩΝ_SLOTS` with exact headers `Therapist | Weekday | StartTime | PatientIDs | Decision`. Every genuine same-therapist/same-weekday/same-clock slot with **more than one distinct patient** must have a single **explicit exact-member approval**: `PatientIDs` is `|`-separated exact confirmed IDs, `Decision=APPROVED_SHARED`. No inferred approvals from source co-location. A new/removed patient, stale approval, duplicate approval, invalid weekday/time, or incomplete membership causes the **entire** admission to stop; all entries are withheld when blockers exist. The existing individual patient/session records remain independent; a shared timeslot must be counted once for provider capacity by downstream logic. Same-patient cross-specialty conflicts remain separately blocked. GitHub commits `49acd07` and `dd0aaa0` include unit scenarios for accepted 2-patient groups, unapproved groups, stale/partial approval, invalid membership. Only synthetic identifiers in tests.

**Scope boundary:** The policy gate is **not** a UI for clinician approval and does not auto-populate or certify 77 existing candidate shared slots. No clinical identity list was added to the repo and no real-data MASTER was written. Latest tests await actual Python run (prior reported 41 green preceded these changes). Native Excel (Windows) save/reopen/rollback and actual shared-slot capacity projection still require dedicated smoke. Avoid calling this complete acceptance.


## 2026-11-12 — Operator confirmed 45 targeted tests passed; unique identity reconciliation started
Operator ran the previous targeted test bundle: **45 passed in 5.99 seconds**. This validates the already committed weekly mapping and exact-member shared-slot approval tests, not native Excel and not the later identity tests.
- Added `Python/rehab_excel/trial_identity_reconciliation.py` (`c7edf16`) with read-only normalized unique-name matching of current source patient rows to old CC IDs. Only unique old-name / unique source-name / unique old-ID combinations yield an **EXACT_NAME_CANDIDATE_UNCONFIRMED**, never an automatically confirmed PatientID. Source clinic+row duplicates, duplicate old names/IDs and duplicate current names are held.
- Added synthetic tests `Tests/test_trial_identity_reconciliation.py` (`636aad1`); not yet run by operator.
- Safety remains: 106 candidate rows, 0 confirmed PatientIDs in private staging. An exact-name candidate is NOT proof of identity. No real patient values may be committed to GitHub. No production/demo .xlsm has been created by this stage.
- Next step: connect the candidate output to a local, reviewable copy of the protected staging workbook, with explicit human approval. Never populate confirmation markers automatically. The post-review gate must still check unresolved cells, patient conflicts and exact shared-slot approvals; native COM importer remains outstanding.


## 2026-11-12 — Date-specific replacement slot protections
- Added pure check `daily_recurring_patient_conflicts` and `validate_daily_replacement_patient_slot` in `Python/rehab_core/patient_schedule.py` (`d6e5774`). Given the effective weekly recurrence and **actual date**, the target appointment is excluded by exact base-entry ID while all other treatments (including those without named providers) block occupied patient hours.
- The core replacement **proposal** `find_replacement_candidates` and final **commit planning** `create_replacement_assignment` in `Python/rehab_core/replacements.py` now accept optional `base_entries` and apply this date-specific filter when supplied (`be74048`); `find_replacement_options` forwards it (`bdd5f89`). Backward-compatible callers without the new argument remain unchanged, so **DO NOT CLAIM ALL EXCEL/VBA ADAPTERS ARE PROTECTED**. Operational callers must supply the actual effective `read_unified_base_schedule` entries, not stale PATIENT_PLANNER cell values.
- Added synthetic regression file `Tests/test_daily_weekly_replacement_slots.py` (`948a8e1`): Monday appointment not blocked by Thursday-only treatment, Thursday alternative blocked by Thursday occupational therapy, final confirmation rechecks, and robotic cross-specialty collision. These tests **await user Windows pytest**; no native Excel COM proof or trial .xlsm write.
- Operational contract: capacity is **distinct therapist weekday/date/clock timeslots**, not patient row count; approved same-slot group members must stay independently identifiable. Replacements cannot consume the patient's other weekday-specific treatments; an absence overlay affects the single date and does not mutate weekly recurrence. Full UI integration still needs verified caller wiring and native Excel smoke.


## 2026-11-12 — Excel global replacement route now checks actual date-specific slots
Operator directed **continuous autonomous development; do not pause for routine acknowledgments**. An audit of the real VBA entry point `ResolveDailyReplacementsInPlace` showed that it invokes `Python/tools/prepare_global_replacements_in_place.py`, which loads `read_unified_base_schedule` and `PROGRAM_WEEKLY_SLOTS`. The global planner has **two** option branches. Wired effective `base_entries` into the non-policy branch in `Python/rehab_core/global_replacement_end_to_end.py` (`d08a2cb`) and into the policy replacement helper (`feab013`) so both now use date-specific patient occupancy. Added an independent, fail-closed **post-global-plan/pre-TSV** verification of each chosen replacement against the exact target date and effective weekly base entries (`f73ec1f`), before VBA can persist changes. Added synthetic regression tests for different weekday/same-treatment and cross-specialty collisions and source-forwarding contract (`1aaa88b`).

**Important limits:** The latest Python changes have not yet passed Windows pytest/native Excel VBA smoke, and standalone manual review/double-booking confirmation paths still require separate call-site audit. Passing earlier 45 tests does not prove these latest commits. Neither CC nor live MASTER was modified. Continue auditing final manual-apply and rollback paths, and never silently infer time from one legacy planner cell.


## 2026-11-12 — Continued without checkpoints: global, manual and final persist slots
Implemented without further operator signoff:
- `d08a2cb`: forwarded actual effective weekly entries to unregulated global replacement proposals.
- `feab013`: forwarded them into policy-governed global replacement suggestions.
- `f73ec1f`: post-planner, pre-VBA TSV patient-date-time validation in `Python/tools/prepare_global_replacements_in_place.py`; silently mismatched dates or missing target sessions hard-stop.
- `1aaa88b`: tests for date-specific global slot protection and wiring.
- `e94603f` and `a5541ce`: clinician-confirmed double-booking exception cannot override another weekday-specific patient treatment, even if the therapist double-booking itself is authorized.
- `c4379be` and `da5c1a1`: final authoritative replacement-fact persistence optionally revalidates effective weekly slots; the real `apply_global_daily_replacements.py` caller supplies them. Existing synthetic/back-compat calls without that argument continue unchanged.
- `87a1dc4`: regression tests prove conflicting Wednesday occupational slot blocks the writer before write; Thursday-only occupancy does not.
No real patient data committed; **latest additions need Python regression and Windows/native Excel smoke** before any xlsm write or operational claims. This report alone is not a green test result.


## 2026-11-12 — 36 daily-replacement regression tests passed; local-only ID review builder added
Operator ran the five requested targeted files: **36 passed in 4.22s**. All of the preceding date-specific patient-slot proposal, global placement, manual double-booking and persistence tests are green in this Windows run; this is still not native Excel UI confirmation.

Continued directly: `trial_identity_review.py` (`f21f63e`) creates **only a NEW local .xlsx review copy**, using the staging `ΑΣΘΕΝΕΙΣ` source identity and the old CC `PATIENTS` registry. The conservative `suggest_identity_matches` results populate **G:H** (suggested old ID/state); **E:F remain unmodified and UNCONFIRMED**. The review builder refuses existing output, wrong file types, used G:H headers, missing source, malformed patient row, duplicate/conflicting old IDs as auto-certification; inputs are unchanged. `c662e81` checks both input hashes and copy before editing.
`Python/tools/prepare_trial_identity_review.py` (`8090001`) is the local CLI; only counts are printed, no names or identifiers. `Tests/test_trial_identity_review.py` (`408ac67`, `b829a77`) tests source/old unchanged, no overwrite, no approval, and an integration check that suggestions cannot bypass `trial_admission_gate`.
**These newly added tests are not yet run.** No private filled review generated in GitHub, and no real .xlsm modified. The actual trial remains hard-stopped until source ambiguities, explicit IDs and shared slot membership approvals are reconciled; do not turn candidate name matches into confirmed hospital identities.


## 2026-11-12 — Actually generated private identity candidate review
Using only locally available private staging and the earlier safety mapping report, created a *separate* private XLSX review: `DIOMIL_ΕΛΕΓΧΟΣ_ΤΑΥΤΟΤΗΤΩΝ_ΠΡΟΤΑΣΕΙΣ_20261112.xlsx`. Exactly **106** current inpatients, **33 unique exact-name ID proposals**, **73 not proposed**, **zero** IDs entered into authoritative E column or approved in F. G:H hold informational suggestions; all other staging review sheets remain present; input files checked unchanged by SHA256. Deliver privately only, **never upload to GitHub**. This was not produced by native Windows COM or a completed master import.
Actual staging column D header was found to be `Θάλαμος/ιατρός (πηγή)`, not `Θάλαμος`: fixed CLI builder expected header in commit `7d6f76a` and test fixture in `78b2aa9`. The Python builder/new synthetic tests await regression execution; the private candidate file itself has been directly exported and read back through artifact_tool.
The 61 source review rows break down as: **29 NO_TIME**, **16 UNRECOGNIZED_TIME**, **11 NO_EXPLICIT_DAY**, **4 UNRECOGNIZED_DAY**, **1 DUPLICATE_DAY**. These categories remain unresolved and must not be implicitly imported. The 16 unrecognized times deserve targeted native Excel time normalization inspection, not auto-fixing. **No clinical workbook import authorized** until patient identity and schedule holds resolved.


## 2026-11-12 — Stable inpatient identity across room, doctor and therapist changes
Operator explicitly clarified that an inpatient **may move room, change responsible doctor and change therapists while remaining admitted**. These are operational changes, NOT discharge/readmission and NOT grounds for creating a second PatientID. Ward, doctor, therapist and the row position in the latest source are **not identity keys**. The clinic/source row is used only as a **snapshot provenance locator** in the import-review staging; a later snapshot can move the same person to another clinic/row without breaking identity. Unique name matching remains only an **unconfirmed candidate**. Hospital MRN should be preferred when independently verified; patient identity must never be inferred merely from location, treating therapist or name.

Implemented pure date-effective primitives:
- `Python/rehab_core/inpatient_placement.py` (`7176a6b`) accepts a **stable PatientID**, baseline patient details, effective date and new room and/or responsible doctor; separate fields can change on distinct dates; historical date queries derive prior values without mutating records; wrong patient, backdating, duplicate same-date events and outpatient ward changes fail closed. Tests `Tests/test_inpatient_placement.py` (`85ce989`).
- `Python/rehab_core/recurring_therapy_changes.py` (`f73f141`) models separately **permanent effective-dated** therapist/time/weekday changes for an exact recurring assignment, preserving the base ID and previous-day materialization. This is independent of **one-day replacement overlays**; it does not silently reschedule multiple weekday-specific entries together. Tests `Tests/test_recurring_therapy_changes.py` (`8d15ef7`).
- `Tests/test_trial_identity_reconciliation.py` (`cbf6126`) explicitly checks a change of clinic/source row cannot invalidate the unique candidate's stable old ID.
**New tests have NOT yet been executed.** These are pure rules only; Excel registry persistence, point-in-time placement history snapshots, UI actions, patient/therapist capacity checks for permanent therapy changes, and controlled COM read-back remain to be integrated. Do not claim earlier finalized daily reports contain historical room/doctor unless the actual snapshot has such fields.


## 2026-11-12 — Permanent therapy-change preflight, no operator checkpoint
Without requiring an extra "OK", implemented `prepare_validated_recurring_change` in `Python/rehab_core/recurring_therapy_changes.py` (`4246128`, `21e764e`). This is a **pure proposal guard**, not a write/apply operation:
1. Project all current effective-dated changes and the proposed edit at its effective date while retaining the immutable baseline recurrence and PatientID.
2. Fail closed on any same-patient/same-weekday/same-time treatment collision.
3. For named providers, require an existing therapist and reuse `check_permanent_assignment` across **every** affected weekday for provider collision and distinct timeslot capacity. Unapproved same-slot group changes need explicit reassessment and are not silently admitted.
4. Require a therapist on ΦΘ/Ρομποτικό; keep provider-optional specialties independent. Return an explicit proposed projection with affected weekdays.
`Tests/test_recurring_therapy_changes.py` updated (`44bc34d`) with five synthetic safe/blocking scenarios. **Not run yet.** This does NOT replace the final Excel native COM transaction or clinician approval.

Also fixed a privacy/safety race in `trial_identity_review.py` (`ef61672`): if an unrelated process exclusively creates the proposed output between the existence check and the atomic open, our error handler must **not delete that other process's file**. Added synthetic race regression (`f9106e4`). New identity review tests also await execution.


## 2026-11-12 — Automated complete Python regression audit
The first expanded CI audit (`3501485`) ran the **entire** Tests/ tree rather than only 79 synthetic safety tests. Ubuntu stopped at 10 reported failures after **714 passes** (12 suppressed warnings, 6.78 s). Windows also reported complete-suite failure. These are **not** failures of the new 79-test safety subset, which remained green. Failure groups include: a stale DAILY_INPUT cancellation options assertion; patient history UI fixture missing the newly mandatory display name; patient and therapist dashboard VBA tests asserting older page layouts/modes; and the native restore repair script expecting a declaration marker no longer found. The last group requires a real code audit before any native write. Do not retroactively declare the whole suite green or silently disable failures.
- `eab7ad9` aligned one legacy test with three explicit daily cancellation/suspension statuses.
- `5ee6781` updated a synthetic patient history fixture with required registry names and the current TSV schema, without weakening the operational fail-closed name checks.
- `9f19a5e` split GitHub Actions into a separate **targeted safety gate** and a **complete legacy+current regression job** for independent, visible pass/fail status on Windows and Linux. The full job is intentionally NOT suppressed; its remaining failures are actionable backlog.
No production workbook or patient data is pushed to CI.


## 2026-11-12 — Complete-suite inventory after green safety gate
The split CI run on Linux covered the entire current test tree: **1548 passed, 20 failed, 14 warnings**. The separate 79-case synthetic safety gate remained green on Linux and Windows. Full-suite failures grouped as:
- H7 patient-history VBA display/layout: 2
- H8 therapist-history mode/chart/labels: 4
- native restore repair marker drift: 5 (do NOT patch production VBA blindly)
- registration/global replacement action stage contract: 1
- therapist-history VBA navigation/protection: 3
- synthetic ReplacementProviderOption fixtures omitted now-required `active_timeslots`: 5
The five fixture errors are test setup drift only and were corrected in `Tests/test_patient_schedule.py` (`f7a006a`) and `Tests/test_replacement_workflow.py` (`9e30fed`); this does NOT change the distinct-timeslot capacity engine. Newly queued CI will verify the impact. Dashboard and restore behavior remain red until verified; do not suppress, skip or force-pass these tests. The prior DAILY_INPUT status and patient history TSV fixture mismatches were reconciled based on current explicit contracts. No clinical source data committed.


## 2026-11-12 — Combined private-source preflight for separate native Excel pilot
Implemented `Python/tools/preflight_trial_excel_sources.py` (`240bccf`) and `Tests/test_preflight_trial_excel_sources.py` (`eeddcd4`); synthetic CI includes this new gate (`98b600a`). The tool joins two read-only checks: (a) `trial_admission_gate.evaluate_trial_admission` on the private review XLSX, and (b) `audit_pilot_source_readiness.inspect_source` on the separately protected CC XLSM. Both inputs are SHA256 verified unchanged after inspection; only aggregate nonidentifying counters and safe issue codes may be printed; no workbooks are saved, patched or imported.

Use from the repository's private Windows machine with institution-approved local paths, **only after closed Excel**:
```powershell
python Python/tools/preflight_trial_excel_sources.py --review "<private-reviewed-staging.xlsx>" --cc "<protected-trial-source.xlsm>"
```
A nonzero exit or any unconfirmed identity, unresolved source cell, cross-specialty overlap, shared-provider approval issue, incomplete CC registry, prior history, VBA problem or other intake issue **blocks** native trial import. Even exit 0 does NOT authorize production release; native COM trial requires separately verified backup, rollback and reopening. Never upload these private files to GitHub Actions. Windows native Excel is NOT installed on GitHub-hosted runners, so CI is synthetic Python-only.


## 2026-11-12 — Full CI regression restored (Windows + Linux)
The prior full regression had six failures: five assertions against an outdated VBA restore patch marker and one registration-menu assertion assuming a now-replaced per-component verification statement. Commits `ca761db`, `d7080f8` preserve the **current validated restore** path and recognize the canonical character-level strike helper at native read-back; tests updated to follow the actual operator-visibility install sequence and shared VBA component verification loop (`8943c8c`, `85c4e4b`, `db05a99`, `33eaf0a`). The production recovery patch still fails closed for absent required markers; an already-secure canonical module does not need to be destructively re-patched.
Authoritative GitHub Actions run `43305214566` at `33eaf0a`: **full Python regression 1,573 passed** on Ubuntu (16 warnings) and **1,573 passed** on Windows (15 warnings). Targeted synthetic safety: **109 passed** on each OS (4 warnings each). These are **Python and VBA source-contract tests**, not native Microsoft Excel COM execution or pilot data import. No real workbooks were modified by this CI run.


## 2026-11-12 — Native Excel isolated read-only probe (not live import)
Added `Python/tools/probe_native_excel_readonly.py` (`4d76049`) and fully synthetic COM-mock tests `Tests/test_probe_native_excel_readonly.py` (`37bb194`). The probe accepts `--isolated-copy <trial-copy.xlsm>` **only for an already isolated, institution-approved non-live copy**; it checks the XLSM VBA project, starts a new native Excel process on Windows, sets `AutomationSecurity=3` (macros forcibly disabled) *before* opening, disables events/alerts/updates, opens read-only with no link updates, verifies required four operational sheets by name only, closes `SaveChanges=False`, quits Excel and verifies the source SHA256 stayed unchanged. It never reads patient records or calls a VBA macro. No save/commit/replacement/import occurs.
Synthetic regression guards reject missing VBA, writable Excel session, missing sheets and external modification; they check no-save behavior and fail-closed macro-security setup.
This is **a structural proof only**. Running it on GitHub-hosted Windows runners with mock COM does *not* demonstrate that Microsoft Excel itself is installed there. Real Excel COM validation still requires a deliberate local Windows execution.
For local operators, *after closing Excel and choosing a detached copy*:
```powershell
python Python/tools/probe_native_excel_readonly.py --isolated-copy "<approved-detached-trial-copy.xlsm>"
```
Stop immediately on a nonzero exit or SHA mismatch. No permission to write the original CC, production MASTER or real clinical roster is implied.


## 2026-11-12 — Read-only workbook recurring-change bridge
Following operator-confirmed synthetic CC native save/reopen A proof (`Documentation/NATIVE_CC_SYNTHETIC_SAVE_REOPEN_20261112.md`), added `Python/rehab_excel/recurring_change_preview.py` (`e7e79c2`). It joins the actual `PATIENTS` registry, unified `PATIENT_PLANNER` + `OUTPATIENT_SCHEDULE` sources, optional granular `PROGRAM_WEEKLY_SLOTS` overlay, and the pure `prepare_validated_recurring_change` capacity/collision gate. The optional granular source REPLACES overlapping legacy (patient,treatment) pairs, never adds duplicate treatment slots. Callers MUST supply approved per-provider capacity for **every** registered therapist, and an exact previous `BaseScheduleEntry` fingerprint; 4-slot therapists do not inherit an assumed 6-slot default. Any stale row shift, unknown provider, invalid capacity, or changed source hash stops the preview.

Tests in `Tests/test_recurring_change_preview.py` (`9f1d63e`, `8b8db11`) cover accepted change, stale row-based identity, unverified/invalid capacity, a 4-slot provider, mid-inspection mutation, and granular weekly overlay. The cross-platform CI targeted suite includes this test (`4705e75`).

**Critical scope:** this is a **read-only proposal**, NOT operational UI, committed event or Excel write-back. The current `weekly:N` and `outpatient:N` IDs derive from Excel row numbers and MUST NOT be persisted as permanent change identity until the workbook schema supplies immutable IDs and authoritative effective-dated events. Existing finalized history must never be retroactively rewritten. Before enabling an actual persistent action: design stable assignment IDs, event history, exact-hash/transaction protection and native save/reopen/rollback tests on a fresh synthetic-only candidate.


## 2026-11-12 — Optional immutable recurring AssignmentID contract
Added read-only `Python/rehab_excel/stable_recurring_ids.py` (`5aaaaa7`, `25b4d18`), integration into `recurring_change_preview.preview_recurring_change_from_excel(..., require_stable_identity=True)` (`88123ae`), regression suite (`ced8383`), and synthetic CI coverage (`73469d7`).

The **opt-in** identity contract is intentionally not installed/written into any existing workbook. For legacy inpatient `PATIENT_PLANNER`, an optional `RECURRING_ASSIGNMENT_IDS` sheet must have exactly the first three column headers `AssignmentID | PatientID | Treatment`: a unique stable patient-and-treatment pair resolves the historical planner assignment regardless of its row. For `PROGRAM_WEEKLY_SLOTS` and `OUTPATIENT_SCHEDULE`, add an explicit `AssignmentID` column **after existing columns**, one ID per source treatment row, so insertions or clock/provider changes do not change identity. Assignments require a unique `assignment:` token matching the strict ASCII identifier grammar in code. Never derive it from a mutable Excel row number, name, therapist, hour or weekday. Wrong patient/treatment at a row is a hard stop.

Missing, duplicated or malformed identity in strict mode is a hard stop. Legacy preview default remains read-only and backward-compatible, but must NOT become a committed effective-dated event. `weekly:N`, `outpatient:N`, `planner:N:treatment` remain non-authoritative transient keys. Existing clinical workbooks are not migrated or overwritten.

**Pending:** unique ID assignment needs a controlled one-time schema migration and independent persisted read-back on a distinct synthetic workbook; event-store and native Excel COM transaction with rollback and actual UI still **not implemented**. None of this preview work proves that a real permanent treatment change has been saved or accepted.


## 2026-11-12 — Synthetic native recurring AssignmentID installer, awaiting local Excel acceptance

Commits `27db1e5`, `e9af67f` implement and test a pure, in-memory migration plan. `ce1d784`, `bd690e2` implement one-shot native COM installation and synthetic-only preflight mocks. Cross-platform CI integration `e9d27b8` resulted in GitHub Actions run `43311751430`: **1,631/1,631 full tests** passed on each of Windows and Linux; targeted synthetic suites **167/167** on each. These results are tests with mocked COM, **not** live Microsoft Excel acceptance.

The new installer `Python/tools/install_synthetic_recurring_ids_native.py` accepts ONLY the frozen `UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm` and a fresh, never-before-created synthetic `_DEMO_` destination. Before a single Excel Save on the COPY only, it verifies source H8 synthetic history, hashes source, generates opaque IDs, verifies ownership and free cells, refuses existing ID registry and occupied G metadata. After Save it independently reads source schedule and assigned stable IDs, verifies H8 history and VBA and unchanged source. No patient therapy is reprogrammed, no event is committed, no production source is touched. Any failure freezes the new candidate and must not be retried with the same path.

Authoritative native PowerShell operator procedure:
`Documentation/NATIVE_SYNTHETIC_RECURRING_ID_MIGRATION_20261112.md`.
**Native migration result remains UNTESTED until the operator posts real PC output.**


## 2026-11-12 — Native 179-ID acceptance and read-only effective-dated event source

Operator confirmed 42/42 focused Python tests and green native Excel COM
migration to fresh isolated synthetic candidate A:
`UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm`.
179 immutable recurring AssignmentIDs were saved and independently read back;
original CC unchanged; 25 synthetic days, 522 patient-session history rows,
150 therapist-day rows, 150 attendance facts and 25 finalization events
remained consistent. Candidate A is **FROZEN**, not a test mutation target.
Evidence recorded in
`Documentation/NATIVE_SYNTHETIC_RECURRING_ID_MIGRATION_20261112.md`
at commit `78de466`.

Implemented optional **read-only** schema reader
`Python/rehab_excel/recurring_change_event_source.py`
(commit `57e35b9`): `RECURRING_CHANGE_EVENTS` sheet, exact headers,
stable AssignmentID ownership, unique opaque EventID, strictly ascending
per-assignment effective dates, duplicate-date refusal, changed-field
validation and historical date projection via existing core.
Tests `Tests/test_recurring_change_event_source.py` cover old/new dates,
multiple events, malformed and duplicated identities, chronology and
non-mutating XLSX readback (commits `7bc30db`, `94a5941`).
GitHub Actions run `43313086923` green on Linux and Windows:
**1,646/1,646 full Python tests** on each platform; targeted synthetic
Windows **182/182**. These are Python tests, not a native Excel event-write
proof.

**Boundary:** this adds no event sheet to the frozen candidate, performs
no workbook write, provides no VBA action/rollback, and cannot approve a
permanent clinical treatment change. Next implement a fresh synthetic-only
event sheet installation, source-bound native transaction with independent
readback and rollback, plus capacity guard over the entire effective schedule.
