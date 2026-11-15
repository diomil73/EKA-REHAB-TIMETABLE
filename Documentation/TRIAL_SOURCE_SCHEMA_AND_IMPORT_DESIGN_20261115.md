# DIOMIL — Real-source demo schema and import design

Date: 2026-11-15  
Branch: `work/afternoon-medical-scope-20261115`  
Status: source schema inspected locally; design only; no patient import or workbook write authorized.

## 1. Product target

The target is one functional, integrated Excel demo based on the operator-reviewed patient/treatment programme, using existing proven Diomil operations. A synthetic workbook is a test fixture and engineering gate, not a substitute for the requested real-source demo.

All real source/review workbooks remain on institution-controlled storage. Never commit workbook bytes, patient names, IDs, room/doctor values, patient-level logs, or private paths to GitHub or public CI.

## 2. Source programme field map

The operator supplied the A:T mapping, and the inspected programme sheet contains additional U/V columns. Preserve the complete source width; do not silently discard U/V.

| Column | Meaning | Import interpretation |
|---|---|---|
| A | Clinic / ward / doctor context | Preserve source context and provenance; blank cells are not automatically forward-filled. |
| B | Patient name | Display identity only; never use name alone as PatientID. |
| C | ΦΘ time | Recurring physiotherapy start time. |
| D | ΦΘ days | Explicit weekday pattern for ΦΘ. Blank-day semantics must be confirmed before parser admission. |
| E | ΦΘ therapist | Provider associated with ΦΘ. |
| F | Reclining time only | A ΦΘ-related operation at a separate time, 30 minutes. It is not a separate ordinary specialty. Its recurrence/day association and replacement behavior must be confirmed before writeback. |
| G | Ε/Θ time | Occupational therapy start time. |
| H | Ε/Θ days | Occupational therapy weekdays. |
| I | Ε/Θ therapist | Occupational therapist. |
| J | Λ/Θ time | Speech therapy start time. |
| K | Λ/Θ days | Speech therapy weekdays. |
| L | Λ/Θ therapist | Speech therapist. |
| M | Ψ/Λ time | Psychology start time. |
| N | Ψ/Λ days | Psychology weekdays. |
| O | Ψ/Λ therapist | Psychologist. |
| P | ΕΦΑ hours | Special physical education time(s). |
| Q | ΕΦΑ days | Special physical education weekdays. |
| R | Robotics time | Robotic treatment start time(s). |
| S | Robotics days | Robotic treatment weekdays. |
| T | Robotics therapist | Robotic therapist. |
| U | Afternoon | Additional afternoon source field; not yet mapped to an accepted operational model. |
| V | Pool | Additional pool source field; not yet mapped to an accepted operational model. |

The source includes Excel numeric and text time representations, grouped weekday patterns, slash-separated values and incomplete/notes-only cells. Reuse the tested `rehab_core.programme_mapping` semantics where they match the source. Any source shape outside that contract must become a review hold, not a guessed value.

## 3. What the existing trial pipeline does

- `trial_identity_review.py` creates a separate identity-suggestion copy. Suggested old IDs are explicitly unconfirmed and do not populate the authoritative PatientID/confirmation columns.
- `trial_admission_gate.py` consumes the normalized review workbook. It requires explicit PatientID confirmation, rejects duplicate IDs, unresolved source holds, cross-specialty overlaps, invalid weekly entries and unapproved shared therapist/day/time membership.
- `preflight_trial_excel_sources.py` is read-only and compares a review XLSX with a separate eligible XLSM intake candidate. It does not import.
- `trial_import_plan.py` is a read-only plan over approved weekly entries and binds hashes of the review and intake workbooks. It does not yet bind the original A:V source workbook.
- `trial_weekly_projection.py` projects only the weekly-slot sheet. It does not project the complete patient registry, demographics/context, all source columns, afternoon/pool fields, or a full operational MASTER programme.
- `weekly_slots_source.py` reads the optional six-column `PROGRAM_WEEKLY_SLOTS` schema and overlays granular entries over legacy patient/treatment pairs. It is not a complete source importer.
- `BaseScheduleEntry` currently has no explicit duration field. The 30-minute reclining operation must not be represented as an independent durationless generic slot if downstream interval/capacity logic needs its duration.

Therefore, the existing source-only tools are necessary admission components, not a complete importer or whole-workbook integration.

## 4. Required target architecture

1. **Private source alignment:** read the original A:V programme and normalized review workbook read-only; verify source row/name/context provenance and every source-cell reference; bind both source hashes. Never print identities.
2. **Identity admission:** explicit operator-confirmed PatientID for every admitted patient; suggested IDs remain suggestions until confirmed. Do not auto-match by name.
3. **Clinical/source review:** unresolved notes, malformed times/days, patient overlaps and shared-provider slots remain held until explicitly adjudicated. The current strict gate requires zero unresolved holds/overlaps; no partial admission is silently allowed.
4. **Canonical patient projection:** create complete Patient registry records, including clinic/ward/doctor context and inpatient/outpatient classification only where supported by explicit source facts.
5. **Canonical recurring schedule projection:** preserve each treatment/day/time/provider as an individual source-bound entry. Do not collapse different weekdays or multiple clock groups into one lossy patient cell.
6. **Additional treatment streams:** model reclining as the specified 30-minute ΦΘ-related operation; map afternoon and pool only after their operational rules and source semantics are confirmed.
7. **Synthetic end-to-end fixture:** use synthetic patient/provider values to exercise A:V parsing, identity gates, holds, weekly projection, patient registry projection, shared slots, collisions and source-hash checks.
8. **Native integration:** only after source/tests are green, build a fresh unique XLSM from a separately eligible trusted source. Use native Excel for save/close/reopen and independent readback. Never modify frozen H10 proof artifacts.
9. **Acceptance:** verify patient identity, every treatment/day/time/provider, source traceability, no phantom sessions, history preservation, navigation and daily workflows. Record exact tested scope and limitations.

## 5. Current private-input observations (non-identifying)

The private review workbook is a proposal/review artifact, not an approval ledger. The patient identity confirmation field is still pending; exact-name suggestions remain unconfirmed; source notes and overlap holds remain unresolved. The review patient rows were cross-checked against the source sheet by source row and patient name, and the two non-patient section/header rows are not patient records. This limited provenance check does not validate every proposed time/day/provider interpretation.

The source context column A is blank on some patient rows, and the review workbook preserves those blanks. Do not forward-fill ward/doctor context without an explicit rule. The source's two clinic sections are represented separately in the review workbook; do not infer clinic from the ward/doctor text.

## 6. Explicit unresolved operator rules

- What does a blank weekday cell mean for each treatment: every weekday, no scheduled days, or an incomplete source hold? The current review proposals expand some blank day fields to weekdays, but this is not accepted as a business rule until confirmed.
- Are U (afternoon) and V (pool) in scope for the first integrated demo, and what are their exact time/day/provider semantics?
- At the implementation point for reclining replacements, request the operator's fresh full instructions. Current known boundary: reclining is ΦΘ-related, separate time, 30 minutes; replacement time is evaluated as an interval and proximity to ΦΘ is preferred when feasible. Do not implement the optimizer from this abbreviated note alone.
- Confirm how missing A context should be handled; never silently inherit a prior ward/doctor value.

## 7. Safety and evidence boundary

No real patient row has been admitted by this design document. No real-data importer, writable destination, native Excel Save/reopen proof or integrated real-data demo has been produced. A positive source-only preflight is not a clinical approval or native acceptance. Keep source/review files, identity decisions, hashes and patient-level evidence private.


## 2026-11-15 — source alignment tooling added and CI accepted

The read-only provenance audit is now implemented in `trial_source_alignment.py`, with private local CLI `audit_trial_source_alignment.py` and synthetic tests. It compares source row/name/context and weekly source-cell references without printing identities or source paths. GitHub Actions run `43751221120` passed 200 focused tests on both operating systems and 1,869 full regression tests on both operating systems. This is source/test evidence only, not native Excel or import acceptance.

The uploaded private source/review pair was separately cross-checked without exposing patient values: source-row/name/context and staged source-cell references aligned; blank A context remains blank. Identity approval and all source/conflict holds remain unresolved. The next implementation should bind the original programme hash into the import plan and add complete patient-registry projection before any native writer.
