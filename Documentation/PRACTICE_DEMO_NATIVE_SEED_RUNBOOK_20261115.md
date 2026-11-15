# DIOMIL practice DEMO: safe native seeding and operational acceptance

Date: 2026-11-15
Branch: `work/afternoon-medical-scope-20261115`
**Current status: code bridge + synthetic tests only. No operator/native Excel acceptance yet.**

## Goal

Deliver **one separate macro-enabled workbook**, with the existing Diomil navigation and daily operational workflows. The first practical acceptance scenario must include initial patients/schedule, new inpatient admission, discharge, patient absence, therapist absence and replacement/restore, new therapist, new student, and save/reopen day. Analytics/H9/H10 visual refinements and fifteen-day progress may follow.

The current source/projection/native-seed modules form a necessary import foundation, but alone DO NOT certify that this interactive workbook is ready for exercise.

## Private data safety

The original programme A:V XLSX, the reviewed identities XLSX and the eligible empty XLSM remain on institution-controlled private Windows storage. **Never commit any workbook bytes, patient names, MRNs, room/doctor metadata, source-cell content or real schedules to the public GitHub repository.**

No source or frozen H10 candidate may be overwritten. Always use a fresh `*_DEMO_*.xlsm` output path outside the Git checkout. A failed native seed must not be retried on the same output. Source SHA256 checks are mandatory before and after any copy/COM write.

## Source and admission gates

1. Operator-confirmed A:V interpretation: no weekend treatments. A treatment time with blank days means exactly ΔΕ, ΤΡ, ΤΕ, ΠΕ, ΠΑ. No time means no slot; explicit day patterns preserved exactly.
2. The original source row, clinic label, patient name and ward/doctor context must match the reviewed workbook with unchanged SHA256. Blank column A stays blank, never inherited from the preceding row.
3. The source and review must contain explicit confirmed PatientIDs, resolved holds/overlaps, and exact shared provider slot approvals. Never promote suggestions automatically.
4. `project_trial_demo_sources` checks the original source again and proves the complete equality of treatment/day/time/provider/source-cell tuples before preparing PATIENTS and granular weekly rows. F reclining and U/V afternoon/pool remain **held** until their independent operational semantics are delivered. There must be no silent partial import.
5. Clinical classification, infectious status and room/doctor/status are independently approved; neither the original context text nor Excel formatting authorizes guessing.
6. The `empty_demo_shell=True` plan is allowed only when the existing workbook inspection finds exactly the `EMPTY_OR_MISSING_PATIENT_REGISTRY` issue: no patients, no history, valid VBA and core sheets. The native seed performs an additional stricter check for empty PATIENTS, treatment planner, weekly source, prior students and active operational sheets.

## Engineering entry points (not an operator command yet)

- `rehab_excel.trial_import_plan.prepare_trial_import_plan(review=..., intake=..., programme=..., empty_demo_shell=True)`
- `rehab_excel.trial_demo_source_projection.project_trial_demo_sources(plan, programme=..., review=..., intake=..., clinical_approvals=...)`
- `rehab_excel.trial_demo_native_seed.seed_trial_demo_workbook(plan, programme=..., review=..., intake=..., output=..., clinical_approvals=...)`

`clinical_approvals` is a **private in-memory** mapping keyed by (review clinic, original Excel row), with `ConfirmedClinicalFields` values. The interface is internal and is not yet a one-click operator workflow. There is no approved private data payload or clean source shell currently available through the GitHub-only execution environment.

The Windows COM backend writes only to the isolated private staged XLSM. It writes PatientID/room/name/infectious/status and optional type/MRN/doctor/context columns in PATIENTS; identical-row identity formulas in PATIENT_PLANNER; one source-bound row per scheduled treatment/day in PROGRAM_WEEKLY_SLOTS; and a detailed DEMO_WEEKLY_VIEW for non-compressible timings. It then performs independent readback, reads the weekly schedule through the real scheduling reader, rechecks inputs, and exclusively creates the final output. Native source visuals and VBA are not removed.

## Critical remaining UI gap

Existing MASTER/PATIENT_PLANNER presentation cannot losslessly represent all distinct treatment times on different weekdays in a single cell. The new DEMO_WEEKLY_VIEW shows the detailed programme, but existing central navigation and replacement screens still need integration/real user validation for those granular rows. A successfully seeded candidate **is not** the final accepted interactive DEMO until those workflows read the same appointments and display them correctly.

Likewise, the current F/U/V treatment holds block whole-source admission. They must either be implemented under separately accepted clinical rules or managed as explicitly unresolved records in a designed, visibly partial practice environment, never treated as secretly dropped sessions.

## Native operator acceptance checklist

Use a **new disposable private candidate**, preserve a frozen clean starting copy for repeatable practice, and only then test:
1. Launch, all forms/navigation/toolbar buttons, identity and clinic sections.
2. Original recurring weekday/time/provider schedule and detailed day view agree; no phantom Saturday/Sunday sessions; no double counting of shared slots.
3. Therapist absence triggers candidate replacement choices, capacity/skill restrictions and approval.
4. Patient absence, return and allowed daily restoration do not destroy immutable recurring assignments.
5. Add inpatient, discharge existing patient, add therapist and student with appropriate eligibility/dates; daily updates reflected in all connected screens.
6. Save/close/reopen through native Excel; verify workbook formulas, VBA project, programme, history and unchanged original source.
7. If any action fails, preserve the candidate and logs, do not force repairs on frozen sources or repeat unneeded old test gates.

**Acceptance wording:** A green Python CI run is code-level evidence only. A new native COM write/readback is a separate milestone; complete interactive operator DEMO acceptance is later still.


## 2026-11-15 — Production reader readback added to native DEMO seed

The native seed's independent post-save verification now also calls the production `read_patient_registry(stage)` parser and compares every imported PatientID, display name, room, infectious flag, status, inpatient/outpatient classification, hospital MRN and responsible doctor with the approved in-memory registry projection. This sits alongside the existing `read_base_schedule(stage)` comparison for every granular recurring treatment/day/time/provider. This closes a readback gap: matching worksheet cells alone did not prove that both production source readers could consume the native result.

Code commit: `702db942b5e9416344438f744efa7a6889dcc5e7`. This commit was made after the previously green `2b1c85bd` workflow; **do not claim the new reader-check commit has CI evidence until a run for this exact SHA completes**. No private source workbook, native Excel run, or operational UI acceptance was performed.


## 2026-11-15 — DEMO legacy planner visual mirror, no schedule compression

Delivered `Python/rehab_excel/trial_demo_planner_mirror.py` and focused `Tests/test_trial_demo_planner_mirror.py`. The private native DEMO seeder now writes a `PATIENT_PLANNER` time/day/provider for each admitted (patient, specialty) **only when** all granular weekdays share one clock and provider and no weekday repeats. Full five-day programmes display `Καθ/να`; irregular weekdays are preserved explicitly. The native readback verifies the mirror. Different clocks, providers, or same-day multiple slots remain exclusively in the full `DEMO_WEEKLY_VIEW` and authoritative `PROGRAM_WEEKLY_SLOTS`; they are NOT silently compressed into an incorrect legacy MASTER cell. The existing daily reader overlays granular weekly entries to prevent double-counting. No original clinical workbook or frozen proof artifact was modified.

**Scope:** user-visible compatibility improvement for the initial practice DEMO, not complete navigation/UX acceptance. Python CI for this new commit must be separately confirmed. Next: ensure the native template and registration/discharge/replacement VBA read the same source, then isolated Windows Excel/COM acceptance with a private empty demo shell.
