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

## 2026-11-15 — Isolated patient-edit dry run (no workbook write)

Added `Python/rehab_excel/afternoon_patient_edit_plan.py` and `Tests/test_afternoon_patient_edit_plan.py`.

`prepare_afternoon_patient_edit` produces an immutable **before/after preview** for **exactly one confirmed patient** inside the independent `AFTERNOON_PHYSIO_SLOTS` namespace. It replaces the selected patient's proposed afternoon rows while preserving every other patient's rows and their relative order. It reports added/removed/unchanged counts, is stable on replay, and refuses malformed or duplicate rows. An empty edit that would remove an existing afternoon programme is **blocked unless the operator explicitly consents** (`allow_clear=True`). This is a pure planning contract, not patient UI, native Excel persistence, automatic import, or actual cancellation ledger.

No change was made to `PATIENT_PLANNER`, `MASTER_SCHEDULE`, `DAILY_INPUT` or existing VBA; the established morning treatment label remains exactly `ΦΘ`, not `Πρωινή ΦΘ`. The technical schema's treatment is `ΦΘ` and the **separate sheet** is its programme discriminator. Never add a duplicate `ΦΘ` dictionary key to the current morning treatment form.

**Unresolved clinical/engineering gate:** same-therapist same-time afternoon rows across different patients may represent approved group therapy, so the pure edit plan intentionally does not independently authorize those sessions. Clinical shared-slot membership approval and actual duration-aware collisions must be linked before enabling an operational writer.

### Added local regression gate

```powershell
git pull --ff-only origin work/afternoon-medical-scope-20261115
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: pull failed" }
$env:PYTHONPATH = (Join-Path (Get-Location).Path "Python")
python -m pytest Tests/test_programme_scope_policy.py Tests/test_patient_programme_draft.py Tests/test_afternoon_slots_source.py Tests/test_afternoon_sheet_preview.py Tests/test_afternoon_patient_edit_plan.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: afternoon patient edit source gate failed" }
```

The most recent **operator-confirmed** result remains `60 passed, 1 skipped in 3.29s` from before these two new files were committed. **Do not** count the new tests as passed until locally executed. No native Excel test is implied.

## 2026-11-15 — Exact-member afternoon shared-slot review (new source gate)

Operator-confirmed local tests for previous slice: **72 passed, 1 skipped in 4.06s**. This result predates the following new changes.

- `Python/rehab_excel/afternoon_shared_slot_gate.py`: pure fail-closed shared physiotherapist slot gate. A group is keyed by normalized therapist + weekday + exact start clock. An approval requires the **complete and exact set of patient IDs**. Unapproved groups, duplicate approvals, partial memberships or stale approvals stop the gate. Error strings never include patient IDs. No VBA, Excel, real identities or workbook writes.
- `review_afternoon_edit_plan(plan, approvals)` always inspects the **entire post-edit afternoon schedule**, not only the modified patient. Adding or removing a member invalidates old shared-slot membership approval and forces explicit new review.
- `Tests/test_afternoon_shared_slot_gate.py`: synthetic tests for missing/exact/stale/duplicate approval, independent therapists and weekdays, patient-only duplicate slots, and approval invalidation during edit.

**Critical limitation:** exact start-time membership approval is only a first-stage gate. Different start times with overlapping durations, clinical group capacity, therapist unavailability, therapist assignment rights and actual schedules remain separate admission requirements. No operational writer may rely on this gate alone. In particular, `validate_afternoon_rows` permits same-time therapist sharing for different patients only so this separate review can decide it, **not** because it is already authorized.

No change is made to the existing MASTER label `ΦΘ`, morning planner treatment keys, DAILY_INPUT or patient UI.

### Next focused local check (tests only)

```powershell
git pull --ff-only origin work/afternoon-medical-scope-20261115
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
$env:PYTHONPATH = (Join-Path (Get-Location).Path "Python")
python -m pytest Tests/test_programme_scope_policy.py Tests/test_patient_programme_draft.py Tests/test_afternoon_slots_source.py Tests/test_afternoon_sheet_preview.py Tests/test_afternoon_patient_edit_plan.py Tests/test_afternoon_shared_slot_gate.py -q -rs
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Afternoon source regression tests failed" }
```

The `-rs` option prints why one test was skipped; do not silently infer the reason from the previous summary. Only after this test gate and further source inspection should a newly named disposable native Excel copy be created. The code is still not a functioning afternoon patient form.

## 2026-11-15 — Explicit-duration overlap gate

**Latest operator-confirmed suite (before this new change):** `88 passed, 1 skipped in 3.66s`. The skip reason remains unverified because the reported output did not include a detailed skip line.

New files:
- `Python/rehab_excel/afternoon_interval_gate.py`: **read-only, standalone** interval gate. Every proposed afternoon session must be paired with an explicitly supplied positive within-day end time. It rejects overlap for the same patient (including different therapist) and for the same physiotherapist (including different patients and different start times). Half-open intervals allow one session to begin exactly when another ends.
- `Tests/test_afternoon_interval_gate.py`: synthetic tests for shifted overlap, adjacency, patient/therapist conflicts, explicit durations, group boundaries, stale approvals and exact end-time mapping.
- Same-therapist group sessions with exactly the same start and end are permissible only with the exact patient-membership approval already required by `afternoon_shared_slot_gate`. Different end times are not silently accepted as a group.
- `review_afternoon_timed_edit` checks the **entire after-edit schedule**, requiring explicit end times for *all* slots including unchanged patients; it fails if there are missing or extra timing entries. It does **not** write or approve the edit.
- **No session duration has been guessed or added to the workbook schema.** The currently defined `AFTERNOON_PHYSIO_SLOTS` sheet has only patient/treatment/day/start/provider and cannot yet supply end time. Before native operations, reconcile the user-approved duration policy and introduce a backwards-compatible authoritative duration source, together with doctor/therapist absence and history/admission gates. Until then, the interval gate is a synthetic contract and its existence does not confer production readiness.

The operator's MASTER instruction remains nonnegotiable: display `ΦΘ` exactly, **never** widen the existing cell by displaying `Πρωινή ΦΘ`; no change to the existing MASTER layout.

### Focused source gate after these commits

```powershell
git pull --ff-only origin work/afternoon-medical-scope-20261115
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: pull failed" }
$env:PYTHONPATH = (Join-Path (Get-Location).Path "Python")
python -m pytest Tests/test_programme_scope_policy.py Tests/test_patient_programme_draft.py Tests/test_afternoon_slots_source.py Tests/test_afternoon_sheet_preview.py Tests/test_afternoon_patient_edit_plan.py Tests/test_afternoon_shared_slot_gate.py Tests/test_afternoon_interval_gate.py -q -rs
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: interval policy tests failed" }
```

This gate runs synthetic data only and does not require opening Excel. Never claim full afternoon scheduling integration until an isolated, native Excel install/save/reopen proof has passed.

## Correction — Reuse existing scheduler interval logic (2026-11-15)

The user correctly identified duplicate engineering. A direct source audit confirmed that `Python/rehab_excel/automatic_patient_schedule.py` already implements `_intervals_overlap`, `_patient_conflict`, `_existing_patient_conflicts` and `_joint_conflict_reason`, with `STANDARD_TREATMENT_DURATION_MINUTES = 45` and `RECLINED_TREATMENT_DURATION_MINUTES = 30`. The existing patient-program VBA also contains a conflict dialog and validation integration. **Do not implement or retain a second standalone interval engine.**

The provisional files `Python/rehab_excel/afternoon_interval_gate.py` and `Tests/test_afternoon_interval_gate.py` were immediately deleted from this feature branch. The historical commits remain in git history for audit only; **do not use their code or old test commands**. The prior documentation section titled "Explicit-duration overlap gate" is superseded by this correction, and any claimed implementation there should be treated as withdrawn. No new check in that deleted test file was run on the operator's computer.

The remaining valid work is `afternoon_slots_source.py`, `afternoon_patient_edit_plan.py`, `afternoon_shared_slot_gate.py` and their tests: pure afternoon namespace and explicit same-start group-membership approval, not a replacement for existing overlap checks.

**Integration objective:** reuse the authoritative duration/overlap algorithm from `automatic_patient_schedule.py`, apply it within the afternoon programme's isolated namespace without inadvertently colliding with morning slots, and preserve existing shared-group approval semantics. Prefer factoring a single tested shared helper if appropriate, with regression tests for morning behavior unchanged and afternoon overlap behavior isolated. Before assuming the 45-minute treatment duration applies operationally to every afternoon session, verify that it matches the existing clinic scheduling contract. The existing code is evidence of implemented defaults, not new authorization to alter clinical practice.

The MASTER treatment label must remain exactly `ΦΘ`, without the word `Πρωινή`; no MASTER width changes.

Latest operator-confirmed focused checkpoint: **88 passed, 1 skipped in 3.66s**, predating the temporary interval implementation and its deletion. All native Excel gates remain pending.

## 2026-11-15 — Existing overlap solver reused, not duplicated

The existing native morning scheduler source `Python/rehab_excel/automatic_patient_schedule.py` implements `_intervals_overlap` for established ΦΘ duration. To address staggered afternoon sessions without another algorithm, `Python/rehab_excel/afternoon_shared_slot_gate.py` now provides `review_afternoon_existing_scheduler_overlaps`, which **calls that existing helper directly**. It first requires exact group membership using the existing afternoon approval gate, then blocks different-start overlap involving the same physiotherapist or patient; adjacent slots remain allowed. Exactly simultaneous, approved shared-group slots remain permissible. There are no morning rows in the adapter input, so same-clock morning/afternoon is not a collision.

New synthetic cases were added to `Tests/test_afternoon_shared_slot_gate.py` for staggered, adjacent, exact group, different provider and patient conflict. This is a source-level adapter, **not** a second engine or Excel integration. The previously removed `afternoon_interval_gate.py` remains absent.

**Critical:** the adapter explicitly reuses the existing 45-minute ΦΘ treatment duration. If afternoon session duration varies in actual clinical practice, do not enable operational saving until the duration policy has been approved and integrated with the authoritative scheduler. This read-only check does not grant clinical clearance, capacity approval, native Excel write permission, or patient import permission. Existing form VBA, MASTER columns and all real workbooks remain unchanged.

Next focused check on the operator's Windows system:

```powershell
git pull --ff-only origin work/afternoon-medical-scope-20261115
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
$env:PYTHONPATH = (Join-Path (Get-Location).Path "Python")
python -m pytest Tests/test_programme_scope_policy.py Tests/test_patient_programme_draft.py Tests/test_afternoon_slots_source.py Tests/test_afternoon_sheet_preview.py Tests/test_afternoon_patient_edit_plan.py Tests/test_afternoon_shared_slot_gate.py -q -rs
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Existing scheduler reuse gate failed" }
```

Previously confirmed result before these adapter changes: **88 passed, 1 skipped in 3.66s**. Do not claim new tests passed until they run. Operator does **not** need to send separate `οκ` confirmations to advance from one safe code slice to the next.

## 2026-11-15 — Read-only patient form state: verified separation prior to VBA patch

Latest operator-confirmed focused suite: **93 passed, 1 skipped in 5.34s**. The sole skip was explicitly identified as `Tests/test_afternoon_sheet_preview.py:68: Symlinks unavailable`; no other skip or test failure.

Added:
- `Python/rehab_excel/afternoon_patient_form_state.py`: read-only adapter with `read_afternoon_patient_form_state(workbook_path, patient_id=...)`, `afternoon_form_status`, and `afternoon_form_rows`. Confirms the selected patient exists in the registry, examines only the separate `AFTERNOON_PHYSIO_SLOTS` source, and returns **that patient's** explicit weekday/time/physiotherapist rows without changing any workbook.
- `Tests/test_afternoon_patient_form_state.py`: synthetic-only workbook regression suite for legacy no-sheet, header-only, two patient identities, malformed registry/schedule, fail-closed invalid rows and byte-identical source after reading.

Legacy no-sheet status is **not installed**, not a writable empty schedule. Header-only PM sheet status is installed-but-empty with zero sessions. The adapter deliberately rejects malformed PM sheets rather than hiding invalid appointments. User-facing status messages are in Greek, but contain no real patient names; object `repr` excludes patient IDs. Actual therapist names may be shown only in a selected-patient interface, never committed to the public repository or logs.

**Integration point found:** `frmPatientScheduling` / `PatientProgramRuntime` currently loads morning treatment keys from `Python/tools/patient_program_bridge.py` and saves using `SavePatientProgramAtomic` into `PATIENT_PLANNER`. Do not add duplicate `ΦΘ` to `lstTreatments` or route afternoon rows into this legacy morning saving function. The new read-only adapter is preparation for a distinct patient-facing afternoon section, not yet installed VBA UI and not an approved writer. Existing morning treatment solver and overlap helper must be reused without replacing the already working morning code.

**Next source gate (no Excel installation):**

```powershell
git pull --ff-only origin work/afternoon-medical-scope-20261115
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: pull failed" }
$env:PYTHONPATH = (Join-Path (Get-Location).Path "Python")
python -m pytest Tests/test_programme_scope_policy.py Tests/test_patient_programme_draft.py Tests/test_afternoon_slots_source.py Tests/test_afternoon_sheet_preview.py Tests/test_afternoon_patient_edit_plan.py Tests/test_afternoon_shared_slot_gate.py Tests/test_afternoon_patient_form_state.py -q -rs
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: patient form read-only source tests failed" }
```

Do not build any native XLSM or touch the clinic's existing MASTER until the chosen synthetic source is confirmed suitable for a fresh disposable integration preview and the new UI has its own independent validated save/rollback contract. The MASTER header/cell still reads `ΦΘ` exactly and is not widened.

## 2026-11-15 — Source integration: read-only afternoon panel on existing patient form

Confirmed from operator on Windows: **101 passed, 1 skipped in 4.22s** before this new integration slice. The one skip is the unavailable symlink test, not a scheduling failure.

New integration:
- `Python/tools/patient_program_bridge.py` adds `afternoon-state-vba` as a **separate read-only command**. Response records are `PM_STATUS\t<0|1>\t<message>` and `PM_ROW\t<weekday>\t<HH:MM>\t<provider>`, never the existing morning `TREATMENT`/`CURRENT` protocol. This reuses the established Python launcher.
- `Python/rehab_excel/afternoon_patient_readonly_vba.py` adds an opt-in, idempotent source patch and a read-only `lblAfternoonState` panel to existing `frmPatientScheduling`. It refreshes after normal configuration and displays the selected patient's PM status and at most two explicit day/time/provider rows plus a count hint. It deliberately has no save, delete or booking button.
- `Python/rehab_excel/patient_registration_scheduling_slice.py` installs the panel after existing MASTER/conflict/combobox installers and before VBA verification; it is only exercised when building a **fresh disposable preview**. It must not be run on frozen evidence or production workbooks.
- New tests: `Tests/test_afternoon_patient_bridge.py` for subprocess TSV protocol, synthetic sheet states, read-only bytes and malformed input, plus `Tests/test_afternoon_patient_readonly_vba.py` for patch safety/idempotence/source contract.
- The existing morning `PATIENT_PLANNER` save logic and the MASTER label `ΦΘ` remain untouched. The new panel increases **patient form height only**, not MASTER column widths or layout.

**Important limitation:** this is first-stage read-only UI integration. No PM operational write path exists, and Python/static tests cannot replace a native Microsoft Excel VBA compilation, installation, save/reopen and manual UI smoke test. In particular, forms, macros, control layout and Unicode display remain unproven in native Excel. Do not merge or deploy this branch until the native proof is green.

### Source gate for this slice

```powershell
git pull --ff-only origin work/afternoon-medical-scope-20261115
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
$env:PYTHONPATH = (Join-Path (Get-Location).Path "Python")
python -m pytest Tests/test_programme_scope_policy.py Tests/test_patient_programme_draft.py Tests/test_afternoon_slots_source.py Tests/test_afternoon_sheet_preview.py Tests/test_afternoon_patient_edit_plan.py Tests/test_afternoon_shared_slot_gate.py Tests/test_afternoon_patient_form_state.py Tests/test_afternoon_patient_bridge.py Tests/test_afternoon_patient_readonly_vba.py -q -rs
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Read-only UI integration tests failed" }
```

No user `οκ` requested. Stop only at this actual Windows test or at an unresolved native Excel/source provenance gate.

**Read-only panel display fix:** the existing `frmPatientScheduling.UserForm_Initialize` sets `Me.Height = 585`, which overrides designer height and would clip the panel at y=525. The isolated patch now changes only the **patient form runtime height to 665** (not MASTER widths); `Tests/test_afternoon_patient_readonly_vba.py` asserts this. See commits `4ab0c78` / `24e0464`. This runtime height adjusted to 665 remains unverified in native Excel until disposable preview acceptance.

## 2026-11-15 — Windows Greek console encoding failure and repair

Native *source-regression* run reported by operator:

```text
3 failed, 110 passed, 1 skipped, 3 warnings in 9.98s
SAFETY STOP: Afternoon integration tests failed
```

**Root cause confirmed from traceback:** `Tests/test_afternoon_patient_bridge.py::_run` used `subprocess.run(..., capture_output=True, text=True, encoding="utf-8")` while the Windows Python subprocess inherited a legacy console encoding (bytes such as `0xC1` for Greek letters). Parent decoding as UTF-8 failed in the subprocess reader thread with `UnicodeDecodeError`; `CompletedProcess.stdout` was `None`, which then caused the three test assertions to fail. This is a subprocess **test harness** mismatch, not a workbook write/corruption or VBA compile error. The failed run is **not green**.

Fix: in `Tests/test_afternoon_patient_bridge.py`, pass `env={**os.environ, "PYTHONIOENCODING": "utf-8"}` to the child process while keeping `encoding="utf-8"` for the parent reader. This matches `PatientProgramRuntime.RunProgramBridge`, which already invokes cmd with `PYTHONIOENCODING=utf-8`. Added a regression that simulates a Greek Windows parent `PYTHONIOENCODING=cp1253` and proves child stdout remains UTF-8. Commits: `da91ef4`, `0624a19`.

**No Excel/VBA production file was changed in this repair**; the fixed source is isolated to tests. The expected symlink skip remains documented. Never interpret the earlier 101-pass checkpoint as an endorsement of these newer UI integration tests.

### Next targeted test (Windows)

```powershell
git pull --ff-only origin work/afternoon-medical-scope-20261115
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: pull failed" }
$env:PYTHONPATH = (Join-Path (Get-Location).Path "Python")
python -m pytest Tests/test_afternoon_patient_bridge.py Tests/test_afternoon_patient_readonly_vba.py -q -rs
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: bridge encoding regression still failing" }
```

If green, rerun the nine-file combined focused suite from the preceding section. Do **not** proceed to an actual Excel COM installer on a non-disposable source until both green and native provenance gates are met. If failure persists, retain full traceback and check child stdout byte encoding; no blanket `errors="ignore"` or replacement decoding permitted, which could corrupt Greek UI text silently.

## 2026-11-15 — Ambiguous 101-pass rerun; full integration gate still required

After the encoding repair (`da91ef4`, `0624a19`) the operator reported `101 passed, 1 skipped in 4.03s`. This is green **for the set actually run** but the exact command is unavailable. The previous failed 9-file run counted `110 passed, 3 failed, 1 skipped`. Therefore `101 passed` **cannot be treated as proof that the three formerly failing bridge subprocess tests were included or repaired**. Never silently upgrade this to an all-green integration status.

Additional conservative change: `_verify_patient_program_vba` now also verifies the afternoon form/runtime method markers, the increased runtime form height, and the existence of the **actual VBA UserForm designer control** `lblAfternoonState`, not just strings in Python. Commit `4618e83`; corresponding static test `dc5be67`. This requires a fresh source gate and later actual native disposable Excel proof. Morning scheduling, original workbook and MASTER widths remain unchanged.

### Mandatory combined local check

```powershell
git pull --ff-only origin work/afternoon-medical-scope-20261115
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
$env:PYTHONPATH = (Join-Path (Get-Location).Path "Python")
python -m pytest Tests/test_programme_scope_policy.py Tests/test_patient_programme_draft.py Tests/test_afternoon_slots_source.py Tests/test_afternoon_sheet_preview.py Tests/test_afternoon_patient_edit_plan.py Tests/test_afternoon_shared_slot_gate.py Tests/test_afternoon_patient_form_state.py Tests/test_afternoon_patient_bridge.py Tests/test_afternoon_patient_readonly_vba.py -q -rs
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: afternoon full source gate failed" }
```

The expected number of tests is **greater than 101** if these 9 files are collected, but the authoritative result is the exact pytest output, not an invented target count. The one symlink skip may persist on Windows. Do not start native Excel COM installation until this full gate is green and an explicitly disposable, non-frozen input has been verified.

## 2026-11-15 — Full source suite green; safe native source eligibility

Operator confirmed **115 passed, 1 skipped in 10.99s** across all nine afternoon source regression files after the Greek subprocess encoding correction. The sole skipped symlink test is unavailable under current Windows settings. This is **not** native Excel compilation, macro proof, or clinical deployment authorization.

A source architecture review after that green run identified that `_install_scheduling_ui` in `patient_registration_scheduling_slice.py` still re-installs the entire patient programme module/form, conflict hooks and MASTER renderer, refreshes existing patient programme cells on MASTER, and saves. **Do not run that broad installer solely to add the read-only PM panel to an existing canonical or frozen workbook.** Its risks are broader than the new PM read-only display. The safe path is to verify a truly disposable, fully integrated source and eventually use a separate minimal native panel installer with explicit untouched-sheet readback. The initial `build_afternoon_sheet_preview.py` creates only a header-only technical sheet and is likewise not a complete form/UI installation proof.

Added isolated **read-only structural XLSM preflight**:
- `Python/rehab_excel/afternoon_native_preflight.py` scans only the ZIP workbook roster and presence of `xl/vbaProject.bin` without opening Microsoft Excel, reading patient rows or saving. It checks presence of `PATIENTS`, `PATIENT_PLANNER`, `MASTER_SCHEDULE`, and stops if the PM technical sheet already exists (because this is a **first-install** candidate). It does not verify actual native VBA behavior, origin or macro security.
- `Tests/test_afternoon_native_preflight.py` uses synthetic ZIPs, covers missing essential sheet/project, existing PM sheet, corrupt XML, duplicate roster and confirms input bytes unchanged.
- `Python/tools/preflight_afternoon_native_candidate.py` is the read-only CLI.
- **Do not mistakenly take a structurally green XLSM for a synthetic source**. The preflight cannot determine whether an XLSM contains real patient information or whether it is frozen evidence. Verify source provenance separately; never publish the contents or use frozen H10 CC/other approved evidence as disposable input.

### Next Windows source gate (tests only, no Excel COM)

```powershell
git pull --ff-only origin work/afternoon-medical-scope-20261115
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
$env:PYTHONPATH = (Join-Path (Get-Location).Path "Python")
python -m pytest Tests/test_programme_scope_policy.py Tests/test_patient_programme_draft.py Tests/test_afternoon_slots_source.py Tests/test_afternoon_sheet_preview.py Tests/test_afternoon_patient_edit_plan.py Tests/test_afternoon_shared_slot_gate.py Tests/test_afternoon_patient_form_state.py Tests/test_afternoon_patient_bridge.py Tests/test_afternoon_patient_readonly_vba.py Tests/test_afternoon_native_preflight.py -q -rs
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: afternoon structural preflight tests failed" }
```

For any future disposable native candidate, after provenance is independently confirmed, inspect structure only with `python Python/tools/preflight_afternoon_native_candidate.py --source "<FRESH_DISPOSABLE_SYNTHETIC_SOURCE>.xlsm"`. **Do not run this placeholder path verbatim.** The CLI neither creates nor modifies an XLSM and is not a go-ahead to install. Native preview source identity, no-op morning proof and comprehensive rollback remain mandatory. No operational afternoon saving or medical absence handling is enabled by this slice.

## 2026-11-15 — Minimal isolated native Excel read-only panel preview (not yet run)

Latest operator-confirmed **source** regression gate: **126 passed, 1 skipped in 10.34s** (the one skipped test uses Windows symlinks). This precedes the new minimal-installer code/tests below.

New implementation:
- `Python/rehab_excel/afternoon_minimal_native_installer.py` has `install_afternoon_minimal_preview(source, output)`. It validates the source's ZIP/worksheet structure and different, **nonexistent** destination, verifies source SHA-256, copies with exclusive file creation and opens **only the copy** in a separate native Excel instance. The source is never opened for edit.
- Before mutation, it verifies existing `frmPatientScheduling` and `PatientProgramRuntime` VBA match the known source-level anchors. It adds only `AFTERNOON_PHYSIO_SLOTS` (five headers, zero data rows, `veryHidden`) and installs the independent `lblAfternoonState` read-only panel onto those already-existing components. It does **not** call broad patient-form installers, patient admission, any morning scheduling writer or MASTER-refresh macro.
- It snapshots worksheet names, MASTER first 12 column widths and used-range values of `MASTER_SCHEDULE`, `PATIENT_PLANNER`, `PATIENTS`. It requires existing values and widths unchanged before save, then closes and reopens the output read-only in a second Excel instance to check the same invariants, new sheet content/visibility and exact VBA source changes.
- On any native failure, leave the newly created output intact as diagnostic evidence, **never** rerun with the same output path. The exception output never intentionally exposes patient row values. This snapshot cannot by itself prove preservation of all workbook formatting, charts, controls, data validation, macros outside the two edited VBA components, or clinical correctness. Do not treat a passed preview as deployment proof.
- `Python/tools/build_afternoon_minimal_panel_preview.py` provides explicit source/output CLI and requires `--confirm-synthetic-disposable`. This flag is an acknowledgement, **not** proof that the workbook contains no real patient data. Verify provenance separately.
- `Tests/test_afternoon_minimal_native_installer.py` checks safe path behavior, independent worksheet baseline, refusal outside Windows, lack of broad installer/MASTER calls, readback contract and confirmation flag. The new tests are **not yet operator-confirmed green**.

### Next source test gate, no Excel open

```powershell
git pull --ff-only origin work/afternoon-medical-scope-20261115
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
$env:PYTHONPATH = (Join-Path (Get-Location).Path "Python")
python -m pytest Tests/test_programme_scope_policy.py Tests/test_patient_programme_draft.py Tests/test_afternoon_slots_source.py Tests/test_afternoon_sheet_preview.py Tests/test_afternoon_patient_edit_plan.py Tests/test_afternoon_shared_slot_gate.py Tests/test_afternoon_patient_form_state.py Tests/test_afternoon_patient_bridge.py Tests/test_afternoon_patient_readonly_vba.py Tests/test_afternoon_native_preflight.py Tests/test_afternoon_minimal_native_installer.py -q -rs
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: minimal native installer source tests failed" }
```

### Future native proof, only after provenance verification

Close all Excel sessions. Identify an actually **fresh, integrated, synthetic, disposable XLSM** known to contain the established patient form/runtime already and no real patient records. Do **not** choose frozen H10 CC, original/production clinical XLSM, source real XLSX, or older evidence files. Read-only structure check:

```powershell
python Python/tools/preflight_afternoon_native_candidate.py --source "<FRESH_SYNTHETIC_INTEGRATED>.xlsm"
```

This command is not a provenance guarantee. Only after verifying its path independently and after source tests pass, an authorized native smoke can use:

```powershell
python Python/tools/build_afternoon_minimal_panel_preview.py --source "<FRESH_SYNTHETIC_INTEGRATED>.xlsm" --output "<NEW_NEVER_USED_AFTERNOON_PANEL>.xlsm" --confirm-synthetic-disposable
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: native afternoon proof failed" }
```

The angle-bracket names are **placeholders** and MUST NOT be run verbatim. **Not authorized yet:** a source file has not been identified/verified in this branch. Native Excel save/reopen, VBA compilation and visible form acceptance remain pending. No afternoon write path, attendance migration, medical deferral scope UI or modification to MASTER label `ΦΘ` exists in this slice.

## 2026-11-15 — 133-pass source checkpoint and native installer security hardening

Operator confirmed `133 passed, 1 skipped in 11.38s` after the first minimal-installer source tests. The sole skipped test is the symlink test unavailable on their Windows machine. **This is not native Excel installation proof.**

Subsequent source audit found a fail-open risk: the minimal installer had a wide `except Exception` block covering lookup of `PatientProgramRuntime`, `frmPatientScheduling`, VBA code inspection and the label. That could accidentally misinterpret **missing or inaccessible VBA modules** as a normal absent label and proceed toward mutation. Commit `db170d7` moves module/form lookup, source readback and source patch validation **outside** the missing-label exception guard. Only an error retrieving the optional label can be interpreted as missing; missing component or source mismatch halts the installer. This stricter behavior still needs local/native verification.

The same commit sets `Application.AutomationSecurity = 3` (`msoAutomationSecurityForceDisable`) on both dedicated Excel instances **before** opening the disposable preview for editing and read-only verification, to avoid executing auto-open VBA macros during technical inspection. The separate Excel application sessions are closed after use; this is not a request to change the operator's Excel macro security globally. It is **not** evidence that native VBA compiles or that the form actually opens.

`Tests/test_afternoon_minimal_native_installer.py` now includes checks for fail-closed VBA component access and forced macro disabling before both opens (commit `7e2a83c`). These newly added tests are not yet operator-confirmed green. Neither the source XLSM nor any frozen or real patient workbook has been touched.

### Targeted fresh source regression

```powershell
git pull --ff-only origin work/afternoon-medical-scope-20261115
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
$env:PYTHONPATH = (Join-Path (Get-Location).Path "Python")
python -m pytest Tests/test_afternoon_minimal_native_installer.py Tests/test_afternoon_patient_readonly_vba.py Tests/test_afternoon_patient_bridge.py Tests/test_afternoon_native_preflight.py -q -rs
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: afternoon installer security tests failed" }
```

**No native command yet.** The next actual native test still requires a verified, fresh, integrated **synthetic disposable** workbook. The preflight and `--confirm-synthetic-disposable` flag do not guarantee provenance. Avoid using H10 frozen CC or production source. The intended test is the isolated minimal installer only, never the broad patient scheduling UI reinstall that refreshes MASTER. All current clinical operations remain unchanged.
