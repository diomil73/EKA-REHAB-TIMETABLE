# DIOMIL authoritative operator and developer handoff, 2026-11-12

**READ THIS FIRST AFTER ONBOARDING.** This document consolidates the last verified results and active boundaries. Older dated H7/H8/H10 instructions are history, not commands to repeat. The current repository branch is `work/history-statistics-rebuild`.

## 1. Actual product target and priorities

Diomil is an application-like Excel/VBA + Python rehabilitation workflow, not a collection of spreadsheet tabs. The goal is a reliable patient/therapist timetable, admissions/discharges, outpatient/inpatient schedules, temporary absences, replacements, permanent therapy/provider changes with historical integrity, attendance and accurate management statistics. Normal users must not navigate hidden technical sheets or type internal row IDs. Preserve familiar navigation buttons, hidden Excel headings/gridlines/sheet chrome where intended and macro-free navigation in management-only XLSX. Do not present unattractive or mathematically impossible charts as accepted.

Most urgent product blocker: a real-data demo MASTER cannot be published while real staging identities, unresolved source cells and scheduling/provider conflicts are not approved. User preference remains accurate trial census first, then lifecycle mechanisms. Synthetic work may safely prepare the machinery but must never be described as imported real patients.

## 2. Authoritative successes, exactly evidenced

### Real Windows Microsoft Excel, operator-reported
1. **Read-only isolated CC probe**, operator pasted `ISOLATED XLSM BYTE-EXACT COPY OK`, `NATIVE EXCEL READ-ONLY PROBE OK`, `NATIVE EXCEL READ-ONLY PROBE GREEN`. Four operational sheets found, VBA present, source unchanged. Frozen: `Excel/previews/UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm`. This is NOT a clinical or writeback proof.
2. **Synthetic CC Save -> close -> read-only reopen**, operator pasted `SYNTHETIC NATIVE EXCEL SAVE REOPEN GREEN` and `SYNTHETIC CC NATIVE SAVE, REOPEN AND HISTORY GREEN`; diagnostic new-sheet marker persisted; VBA preserved; source CC unchanged; independent synthetic history matched. Frozen: `Excel/previews/UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm`. This is NOT Apply/Restore.
3. **Stable AssignmentID migration on new synthetic CC copy**, operator pasted `42 passed in 3.06s`, `SYNTHETIC NATIVE ASSIGNMENT IDS GREEN`, `Verified immutable assignments: 179`, `Original CC unchanged: YES`, `Historical synthetic records independently verified: YES`, and final `SYNTHETIC CC ASSIGNMENT ID NATIVE MIGRATION GREEN`. Frozen: `Excel/previews/UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm`. Independently checked 25 synthetic service days, 522 patient-session rows, 150 therapist-day rows, 150 attendance facts, 25 finalization events. This is **not** an event-write test.
4. Earlier historical native Excel smoke acceptance for DAILY_INPUT absence Apply/Restore X and finalized H2 Z is recorded in `Documentation/CURRENT_HANDOFF.md` and specialist runbooks. Do not reuse those frozen candidate filenames. Do not extrapolate those successes to permanent recurring changes.

### GitHub CI
Latest fully observed run before this handoff: `https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/43313086923` at commit `94a5941`: **1,646/1,646 full Python tests on Windows and Linux**; Windows targeted synthetic safety **182/182**. CI executes Python tests and mocked COM contracts, **not actual Microsoft Excel**. Test warnings are not automatically defects or proof of native acceptance. Recheck latest GitHub Actions after every subsequent source change, and report exact job conclusions.

## 3. Frozen and source artifacts

- Trusted synthetic source: `Excel/previews/UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm`. Source only; **never mutate, reinstall or overwrite**.
- Frozen successful read-only probe A, native Save/reopen A and 179-ID migration A as above; **never reopen for editing, delete, overwrite or use as subsequent writable candidate**.
- Prior H7/H8/H10 BO, BU, BV, BW, BX, BY and many older candidates have separate statuses documented in `CURRENT_HANDOFF.md` / `PROJECT_STATUS.md`. Never infer a candidate is reusable from its filename. The presence of a file is not successful evidence.
- `Rehab_Center_System_v27_1.xlsm` is immutable baseline.
- Actual private patient source `ΠΡΟΓΡΑΜΜΑ`, the local staging review XLSX and any operational workbook contain sensitive clinical information. **Never upload them to public GitHub, CI artifacts, public logs or management exports.**

## 4. What code exists today

- `Python/rehab_core/recurring_therapy_changes.py`: in-memory effective-dated assignment changes, collision/capacity proposal checks and preservation of dates before effect.
- `Python/rehab_excel/recurring_change_preview.py`: read-only Excel schedule bridge, optional weekly granular overlay, capacity explicitly provided for EVERY therapist, original file SHA validation, stale-row fingerprint rejection; `require_stable_identity=True` invokes strict stable-ID binding.
- `Python/rehab_excel/stable_recurring_ids.py`: optional immutable IDs reader. Inpatient planner IDs in `RECURRING_ASSIGNMENT_IDS` keyed by PatientID+Treatment; optional `AssignmentID` seventh column of `PROGRAM_WEEKLY_SLOTS` and `OUTPATIENT_SCHEDULE`. Legacy `weekly:N`, `planner:N:treatment`, `outpatient:N` are not valid persistent event IDs.
- `Python/rehab_excel/recurring_id_migration_plan.py`: pure in-memory UUID-based one-time proposal, fail-closed duplicate detection.
- `Python/tools/install_synthetic_recurring_ids_native.py`: **synthetic-only** one-shot native installer from frozen CC to NEVER-EXISTING distinct `_DEMO_` output; preflight, exclusive copy, macro-disabled Excel Save, independent ID/history/VBA readback; never retry a candidate filename.
- `Python/rehab_excel/recurring_change_event_source.py`: optional **read-only** `RECURRING_CHANGE_EVENTS` worksheet parser, unique EventID/AssignmentID, patient binding, strictly ascending effective dates, no same-day duplicate, valid changed fields, projection without historical rewrite.
- `Python/tools/prepare_native_probe_copy.py`, `probe_native_excel_readonly.py`, `prove_synthetic_cc_native_write.py`: completed structural native smoke foundations, NOT functional recurring-change Apply/Restore.
- `Python/rehab_excel/trial_admission_gate.py`, `Python/tools/preflight_trial_excel_sources.py`: **read-only** real-source safety gate; incomplete patient identity, unresolved cells, overlapping specialties and provider groups still block actual import.

No native permanent-event append, authoritative transaction, rollback, operational VBA UI, full history integration or real-patient migration is approved yet.

## 5. Local checks and safe PowerShell conventions

From repository root on Windows, close Excel first. Run a complete PowerShell scriptblock in ONE selection so `throw` aborts subsequent commands:

```powershell
& {
    $ErrorActionPreference = "Stop"
    git pull --ff-only origin work/history-statistics-rebuild
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: git pull" }

    python -m pytest Tests/test_recurring_therapy_changes.py Tests/test_recurring_change_preview.py Tests/test_stable_recurring_ids.py Tests/test_recurring_id_migration_plan.py Tests/test_install_synthetic_recurring_ids_native.py Tests/test_recurring_change_event_source.py -q
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: recurring tests" }

    Write-Host "SOURCE CONTRACTS GREEN" -ForegroundColor Green
}
```

For the entire regression, separately run `python -m pytest Tests -q --disable-warnings --maxfail=35`. It is **not** a native proof. Never use `--overwrite` to force through a one-shot failure; never resume running individual commands after a failed PowerShell block. For a native task, pick a truly unused candidate suffix, check source SHA and contracts, perform at most one operation, close Excel, independently reopen read-only and inspect outcomes, then freeze regardless of success/failure.

**Do not rerun** any of the three successful 2026-11-12 native acceptance runbooks against their already existing output paths. Those scripts are preserved for historical reproduction guidance, not as next steps.

## 6. Current blockers and exact next implementation order

1. Extend the optional event source into a strict **synthetic-only append plan**, rejecting stale source hash, duplicate IDs, backdated or same-day assignment events and unapproved provider capacity. No real data and no workbook write initially.
2. Build a fresh one-shot synthetic Excel native event-sheet installation/append workflow; source-bound exclusive new output, Save/close/reopen, immutable event readback and consistency audit. Do not modify frozen `*_RECURRING_IDS_20261112_A.xlsm`.
3. Add a rollback/reversal design for permanent events (compensating event or versioned cancellation, never delete historical fact), validate past/current/future readback and affected patient/provider capacity. Explicitly test 4-slot versus 6-slot therapist limits, shared approved provider slots and multi-day conflicts.
4. Only then connect application-style VBA form/selection buttons with transaction/rollback and **new** isolated synthetic native operator proof.
5. Separate real-data path: confirm all PatientIDs, resolve/quarantine 61 ambiguous source cells and 11 candidate patient-time overlaps, adjudicate shared provider slots, then verify admission gate on approved private files. Numbers are historical staging diagnostics, NOT current approved census; DO NOT fabricate or auto-approve them.
6. Revisit visually unaccepted H10 management UI/export and true occupied-timeslot capacity diagnostics. Internal management XLSM is not share-safe; independent macro-free whitelisted XLSX requires privacy and visual approval.

## 7. Failure handling and evidentiary discipline

- If a gate fails, stop immediately, preserve failed candidate untouched, record first error, commit a source fix and choose a **new** distinct candidate name. Never try twice against the same file.
- Do not report success from a final banner printed after a previously failed command.
- No native Excel proof may be inferred from GitHub Windows Python runner. No patient facts or names in public repo or logs.
- Check independently: source hash, output identity, VBA project, expected source sheets, readback, historical consistency and separate outcome inspector.
- Do not confuse therapist occupied **timeslots** with total delivered **sessions/patient lines**; a therapist may have >6 sessions via group appointments while no more than approved released clock slots. Do not invent 8 available slots.
- Preserve pre-change service-day history and versioned finalized facts. A new effective-dated change must not edit existing historical snapshots.
- Do not claim a report/export is ready for management until privacy, numeric and presentation gates pass.

## 8. Documentation update contract (user explicitly reaffirmed)

**After every meaningful change**, update all affected living documents in the same work cycle:
- `ONBOARDING.md`: authoritative latest pointer and architecture overview.
- `CURRENT_HANDOFF.md`: exact last green/failed test, commits, frozen files, next safe commands and contingencies.
- `PROJECT_STATUS.md`: achieved versus open work, real-data/import blockers, priorities.
- `IMPLEMENTATION_MAP.md`: module ownership and test locations.
- `DECISIONS.md`: durable rules, data model constraints, failed-closed and never-overwrite principles.
- `CHANGELOG.md`: meaningful functionality/status.
- Relevant `Documentation/NATIVE_*.md` and subsystem docs: scenario, full single-block PowerShell commands, expected evidence, failure/stop and reuse prohibitions.

When these disagree, the newest dated **operator-confirmed evidence** and this handoff supersede older planned instructions; do not silently rewrite old historical observations. Every new context should first read `ONBOARDING.md`, then THIS file, then `CURRENT_HANDOFF.md` and the relevant subsystem runbook.

**Status when this file was created:** 179-ID native migration confirmed; event source READ-ONLY; NO native event append or rollback proven. No new user action requested.


## 9. Subsequent source-only checkpoint: strict recurring-event append plan

After the 1,646-case validated event reader, committed `Python/rehab_excel/recurring_change_append_plan.py` (`a0791a3`) and `Tests/test_recurring_change_append_plan.py` (`3a44486`) with cross-platform CI wiring (`38b3b28`). This is a **read-only** in-memory candidate for one effective-dated event. It verifies immutable AssignmentID, exact expected BaseScheduleEntry, existing event IDs and ledger schema, strict later-than-existing effective date for the same assignment, and invokes `preview_recurring_change_from_excel(..., require_stable_identity=True)` with the actual current event history and explicitly supplied therapist capacities. It records source SHA256 and provides `verify_unchanged_source` for a future transaction. It never creates a workbook, appends an Excel row, saves, executes VBA, approves a clinical change or makes history authoritative.

**Next gate:** inspect full GitHub Actions results for the latest source commit, repair any failed case before proceeding. Then add tests proving source changes while parsing, bad existing event ledger schema/row gaps, same-date collision, actual 4-slot capacity and future effective-date conflict, and only after that implement a NEW synthetic-only unique output native append + independent readback. Frozen A files and original CC may not be changed. This newer checkpoint supersedes the earlier "no append plan" statement; it does **not** supersede any native-success evidence.

**Verified source CI outcome:** GitHub Actions run `43314844755` at commit `38b3b28` passed all four jobs: **1,652/1,652 full Python tests on Windows and Linux**, targeted synthetic safety **188/188 on Ubuntu** (Windows synthetic job green as well). This is Python/mock COM only, no native Excel event append. Documentation-only commits after this run must not be mislabeled as newly native verified.


## 2026-11-13 priority checkpoint: real-source integrated trial demo

The user explicitly reconfirmed that the **next visible milestone is a single usable Microsoft Excel demo built from the recently supplied real clinic schedule**, with already implemented Diomil mechanisms. This real-source demo was in preparation before recent synthetic H10/recurring-change work. Do not restart the architecture or substitute another purely synthetic statistics demo.

**Immediate next steps:** inspect existing private real-source review and read-only admission pipeline (Python/rehab_excel/trial_admission_gate.py, Python/tools/preflight_trial_excel_sources.py, and tests). Find latest actual private staging/preflight evidence and re-evaluate unresolved patient identity, source cells, patient-time conflicts and shared provider slot decisions. The historical figures **61 ambiguous cells and 11 candidate overlaps** are staging observations, not an approved or necessarily current census. Do not silently import, fabricate, auto-approve, publish or leak private patient records.

Identify native-proven application mechanisms and assemble a NEW isolated, never-before-used macro-enabled trial candidate with only verified facts. Newer recurring-event writeback/rollback may remain disabled pending separate native proof. Preserve immutable baseline, private source, synthetic CC and all frozen evidence. Require source hashes, no-overwrite gate, independent native Excel Save/close/reopen/readback, privacy review and operator acceptance. Python CI green does not establish native Excel proof.

**Status:** next priority recorded; **no integrated real-data demo has yet been completed or accepted**. No delivery date was promised. Update all affected docs as implementation progresses. The user will resume next conversation and prefers proactive work without repeated OK.
