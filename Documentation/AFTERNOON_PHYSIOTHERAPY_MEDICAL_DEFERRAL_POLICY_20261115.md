# DIOMIL — Afternoon physiotherapy, pool and medical deferral scope (2026-11-15)

Status: **pure policy and synthetic tests committed on an isolated branch; NOT operationally integrated**.

## Operator-confirmed requirements

1. The 2026-11-12 real workbook is only a source of refreshed patient/treatment facts. Never replicate its layout inside the canonical DIOMIL MASTER. The existing doctor-under-patient display convention remains unchanged.
2. Source column map for the legacy `ΠΡΟΓΡΑΜΜΑ` sheet: A clinic/ward/doctor, B patient, C/D/E morning physiotherapy time/days/provider, F tilt-table time, G/H/I occupational therapy time/day/provider, J/K/L speech therapy, M/N/O psychological therapy, P/Q EFA time/days, R/S/T robotics time/days/provider, U afternoon programme (do not import), V pool time. The actual source positions are unchanged; any future DIOMIL UI layout is independent.
3. Pool is a **morning** treatment. The existing `Πισίνα` weekly treatment code already exists in `trial_admission_gate.py` and `weekly_slots_source.py`. Do not duplicate this taxonomy or guess missing day/provider.
4. Afternoon programme is **physiotherapy only**. Its patient-management fields are `time`, `day`, `physiotherapist`. They start **blank**; do not import unreviewed legacy column U or automatically create any PM session.
5. Morning and afternoon schedules use **independent conflict namespaces**. A shared time, patient or therapist across the two programmes does not count as a cross-programme scheduling collision. Within each programme, reject overlapping sessions sharing a patient or assigned therapist, subject to the existing explicit group-therapy approval contract when integrated.
6. Temporary discharge and transfer are patient-wide states covering both morning and afternoon, for the relevant effective dates. This initial pure policy models **one explicit date** only, not a date range.
7. Medical deferral requires an explicit scope:
   - whole day: all scheduled patient therapies including afternoon;
   - specific time window: only sessions overlapping that actual window (and optionally a selected programme);
   - specific session: exactly the selected session ID, not every physiotherapy session.
   Do not infer duration from start times. Never erase recurrence rows because of a one-day deferral.
8. Morning-specific physiotherapy deferral, or absence of the morning physiotherapist, does not imply cancelling afternoon physiotherapy with its separately assigned therapist. The converse is also true.
9. A user-confirmed cancellation must be previewed with the exact affected sessions before save. Record reason, scope, date, targets, time bounds and restore/amendment history. No silent retrospective rewrite.

## Implementation slice committed

- `Python/rehab_core/programme_scope_policy.py`: pure immutable scheduling and absence records; `affected_session_ids` preview and `conflicting_session_pairs` local namespace checker.
- `Tests/test_programme_scope_policy.py`: **synthetic-only** regression scenarios; no patient names or identifiers from real data.
- `Python/rehab_core/patient_programme_draft.py`: pure patient-management draft contract for the blank afternoon physiotherapy triple; explicit time/day/therapist input required; existing morning pool remains distinct.
- `Tests/test_patient_programme_draft.py`: synthetic-only draft validation and pool placement tests.
- `Python/rehab_excel/afternoon_slots_source.py`: new separate, read-only afternoon source schema `AFTERNOON_PHYSIO_SLOTS`, with headers `PatientID, Treatment, Weekday, StartTime, Therapist`. Only `ΦΘ` is allowed. Empty/header-only represents zero afternoon sessions. It does not install or create a sheet in a real workbook.
- `Tests/test_afternoon_slots_source.py`: synthetic-only source roundtrip and fail-closed tests (invalid patients, missing fields, duplicate same-patient afternoon slots, no inferred afternoon data).
- Both files do **not** modify any Excel workbook, VBA, MASTER, patient registry, CC or frozen native acceptance artifact.
- This is a policy demonstration, **not** installed in DAILY_INPUT, UI, patient management, scheduling persistence or stats. Do not claim end-to-end completion.

## Integration tasks before operational release

1. Read the existing patient manager forms and insertion/edit macro source. Add optional afternoon-physio controls and morning-pool option only where missing, without changing the visual MASTER layout or treatment legacy semantics.
2. Keep PM fields blank for every imported patient; require explicit date/day, clock and therapist before saving an afternoon appointment. Validate the three fields together; do not infer absent values from AM physiotherapy.
3. Bridge operator selections to an established daily-change ledger and reversible day history, after confirming current protected-sheet write paths. Do not overwrite or bypass existing daily-close/version-chain guarantees.
4. Reconcile morning and afternoon duration semantics and group/shared therapist sessions from existing application rules before reusing the pure conflict function for real data.
5. Add two-stage preview and confirmation of affected sessions for full-day or medical window deferral, then exact native Windows Excel smoke on a **fresh uniquely named disposable XLSM**, with independent read-only outcome inspection.
6. Include afternoon-specific statistics without contaminating morning workload, and ensure cancellations record accurate treatment/time/provider outcomes. Keep current daily/weekly/historical data unchanged.
7. All real source staging, identity matching and patient data remain **local-only and private**; repository is public. No raw source, hashes used as source fingerprint if confidential, names, row dumps or patient details in Git/CI.
8. Preserve: production baseline, original 2026-11-12 source, frozen synthetic CC and every accepted/failed/uncertain smoke XLSM. No overwrite, repair-in-place or `git clean`.
9. Perform focused synthetic regression suite, full repository suite, then actual Excel native user acceptance. Tests are not considered green without output evidence.

## Focused local test command (repository root)

```powershell
git fetch origin
git switch work/afternoon-medical-scope-20261115
$env:PYTHONPATH = (Join-Path (Get-Location).Path "Python")
python -m pytest Tests/test_programme_scope_policy.py Tests/test_patient_programme_draft.py Tests/test_afternoon_slots_source.py Tests/test_afternoon_sheet_preview.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: afternoon/medical policy tests failed" }
```

Before switching, run `git status --short` and ensure tracked work is preserved. Untracked local previews/forensic backups must remain untouched. For source state the previous main work branch remains `work/history-statistics-rebuild`. This feature branch must not be merged until coverage and native integration gates are satisfied.

## Troubleshooting / stop conditions

- Missing or ambiguous therapist, time/day or programme: do not create afternoon record.
- Afternoon specialty other than physiotherapy: refuse.
- Patient state full-day vs targeted session: never infer based on free-text reason; choose scope explicitly.
- Target absent or session id belongs to another patient/date/programme: refuse.
- A patient has no afternoon entry: no phantom afternoon cancellation; full-day still cancels valid morning entries.
- Two sessions share a time across AM and PM: **not** a conflict based on cross-programme overlap alone.
- Two sessions overlap within the same programme: apply the existing clinical/shared-session approval policy before operational rejection or admission.
- Historical day is finalized: use amendment/version chain, not direct edit.
- Any native save, restore or read-back failure: stop once, preserve evidence; no second attempt on the same smoke copy.

## Existing patient-manager integration map (read-only source audit, 2026-11-15)

- `Python/rehab_excel/patient_edit_form_vba.py`: edits patient identity/status/ward/doctor and hands off to patient-treatment programme. **Do not** put the afternoon triple into doctor/identity controls.
- `Python/rehab_excel/patient_program_form_vba.py`: `frmPatientScheduling` loads per-treatment choices from `SETTINGS` using `LoadPatientProgramConfig`. It stores selected treatment/day/time/provider in treatment-keyed dictionaries; it is not partitioned by morning/afternoon. Adding a second treatment named `ΦΘ` to that list would collide with the existing morning selection. The SavePatientProgramAtomic path targets `PATIENT_PLANNER` and must NOT be reused to persist PM as a second copy of morning `ΦΘ`.
- `Python/rehab_excel/patient_program_draft.py`: treatment catalogue comes from `settings.therapies`, while provider requirement for the legacy morning `ΦΘ` comes from `_REQUIRED_PROVIDER_TREATMENTS`. This means the existing pool must be checked in `SETTINGS` before claiming it appears in the form; the existing `Πισίνα` render and admission code are not by themselves proof of a visible choice.
- `Python/rehab_excel/patient_program_master_render_vba.py`: confirmed current MASTER treatments include `Πισίνα` written to the established morning pool column (MASTER column 6). The doctor-under-patient presentation must stay as-is.
- `Python/rehab_excel/patient_registration_scheduling_slice.py`: assembles the existing VBA form/runtime and saves the workbook. New PM support should use a distinct optional storage namespace and audited installer integration, not an in-place rewrite of this active stack before native acceptance.
- `Python/rehab_excel/weekly_slots_source.py`: the morning pool is already an admitted treatment token, and this legacy weekly source has no programme discriminator. Adding afternoon `ΦΘ` to its existing six-column rows would lose namespace separation. Do not silently overload.
- The new `patient_programme_draft.py` is a **separate pure model**, not a UI control and not a replacement for `patient_program_draft.py`; their distinct names are intentional. A future integration should explicitly bridge from forms into the policy, not confuse these modules.

### Required next engineering slice

Design and test a distinct, opt-in afternoon programme data store keyed by patient + weekday + time + physiotherapist; ensure zero auto-seeding from the real source and no effects on the existing morning planner, solver, MASTER layout, histories or conflict guard. Then add native form controls only after the source-level interface and rollback contract are audited. Avoid duplicate treatment identifiers by using a programme field/namespace rather than a new duplicate `ΦΘ` treatment name.

This map is based on repository source inspection, not native Windows Excel behavior.

## 2026-11-15 — Native Excel *header-only* isolated sheet preview

- New files: `Python/rehab_excel/afternoon_sheet_preview.py`, `Python/tools/build_afternoon_sheet_preview.py` and `Tests/test_afternoon_sheet_preview.py`.
- The builder takes an explicit source XLSM and **a fresh, nonexisting output XLSM**. It creates `AFTERNOON_PHYSIO_SLOTS` with the five agreed technical headers and **zero patient rows**. The sheet is `veryHidden`; the future patient-facing UI is not yet implemented.
- It never opens the source for editing. It compares source SHA-256 before/after copying and after native Excel save, reopens output read-only, verifies the VBA project, header-only state, technical hidden status and widths of the first 12 MASTER columns.
- It refuses to overwrite any existing output, use a symlink output, use an XLSX instead of XLSM, or install on a file missing MASTER/PATIENTS/PATIENT_PLANNER and VBA.
- It intentionally **preserves a failed output** as potentially valuable forensic evidence. If failure occurs, do not retry with the same output path or manually repair that copy. Use another fresh name only after reviewing the failure.
- This native step does **not** change the existing morning planner, install buttons or forms, make PM scheduling operational, or admit a real patient. It is an infrastructure smoke test, not production clearance.

### Safe local commands

Close Excel. After `git status --short` shows no conflicting tracked local changes, pull the branch and run the synthetic tests:

```powershell
git pull --ff-only origin work/afternoon-medical-scope-20261115
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: pull failed" }
$env:PYTHONPATH = (Join-Path (Get-Location).Path "Python")
python -m pytest Tests/test_programme_scope_policy.py Tests/test_patient_programme_draft.py Tests/test_afternoon_slots_source.py Tests/test_afternoon_sheet_preview.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: source gate failed" }
```

For the native step, choose a **fresh disposable integrated synthetic Diomil build**, not the frozen H10 CC, original clinical source, production master or any existing approved smoke artifact. Copy the source to a new separately named disposable input only if its provenance and native-integrated status are understood. Then:

```powershell
python .\Python\tools\build_afternoon_sheet_preview.py `
  --source ".\Excel\previews\<FRESH_SYNTHETIC_INTEGRATED_INPUT>.xlsm" `
  --output ".\Excel\previews\<FRESH_NEW_AFTERNOON_COPY>.xlsm"
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: native sheet preview failed" }
```

The placeholder paths are intentional: do not execute verbatim or substitute frozen CC automatically. The operator must verify existing files and choose unique names. A successful result demonstrates only the technical sheet and MASTER width preservation. Run separate source/VBA tests and interactive native UI acceptance before making an operational claim.

## Confirmed local Windows regression checkpoint — 2026-11-15

The operator ran the focused synthetic suite on the feature branch and reported:

```text
60 passed, 1 skipped in 3.29s
```

This is **not** native Excel proof, and no workbook was built or opened by this result. One test was skipped. `Tests/test_afternoon_sheet_preview.py::test_refuses_symlink_output` is written to skip if symlink creation is unavailable, so it is a plausible cause, but **the exact skip reason has not been verified**. When useful, run `python -m pytest Tests/test_afternoon_sheet_preview.py -q -rs` without requiring extra user approval for routine source work.

The operator also explicitly required:
- MASTER must continue showing **`ΦΘ`** exactly, **not** `Πρωινή ΦΘ`: no wider column/cell or visual change.
- In patient-management UI, morning and afternoon may have distinct clear labels; they are two programmes for the **same treatment kind** `ΦΘ`. PM must have its own storage, while morning remains mapped to the established `PATIENT_PLANNER` fields.
- Proceed autonomously between actual gates; do not ask the operator to respond `οκ` after each increment. Ask only for local/native Excel tests, true clinical decisions, or unresolved ambiguities. Continue to commit docs/contingencies for each meaningful step.

**Stop before a native build until the exact input is verified as a fresh synthetic full integrated workbook.** The existing `Python/tools/build_unified_preview.py` is not by itself evidence that every later H10 integration is present. Never repurpose frozen H10 CC, X/Z/AB/BF or original source as a disposable target. Keep unrelated changes isolated and do not merge this branch based solely on the 60 synthetic passes.
