# MASTER REQUIREMENTS RECOVERY PLAN

Date: 2026-10-29
Trusted baseline: `4a2eed09ae0073ea61d34d24d55e5e340269d3a6`
Frozen baseline branch: `recovery/known-good-4a2eed0`

## 1. Purpose

This document is the authoritative recovery and implementation plan after reverting the active development context to the last trusted handoff from the previous conversation.

The governing rule is simple: functionality that was already green and working at the trusted baseline is not to be reopened or redesigned without new evidence. New work must be layered on top of the known-good contracts, in isolated branches and disposable workbook copies.

## 2. Safety rules

1. `Rehab_Center_System_v27_1.xlsm` is immutable.
2. `recovery/known-good-4a2eed0` is a frozen reference branch and must not be used for experiments.
3. New work happens only on child branches created from the trusted baseline.
4. Every native Excel smoke uses a newly named `.xlsm` output. Existing known-good previews are not overwritten.
5. No repair-on-repair workflow. If a disposable smoke becomes uncertain, discard it and rebuild a fresh copy from the canonical pipeline.
6. No business-rule changes are made in VBA. Python remains the owner of scheduling, matching, replacement, capacity, and historical business logic.
7. The implementation order is always: rule/contract -> tests -> Python -> green -> Excel integration -> native proof -> documentation checkpoint.
8. No closed subsystem is reopened unless a focused test or native reproduction proves that subsystem is the failing layer.
9. Experimental post-baseline V2 patch files are requirements/reference only. They are not to be copied back blindly.
10. All meaningful behavior changes and native findings must be recorded in the authoritative documentation before moving on.

## 3. Trusted closed areas at the baseline

The following are treated as known-good at `4a2eed0` unless new evidence proves otherwise:

- restart-safe authoritative daily reconstruction
- DAILY_INPUT operational overlay semantics
- daily replacement fact lifecycle already completed by the baseline handoff
- restore lifecycle already completed by the baseline handoff
- daily statistics reconstruction parity for current-day operational state
- outpatient recurring workflow and outpatient create/edit/conflict handling
- therapist workload strip / presentation work that was already green
- MASTER/application-shell work already completed before the history milestone
- confirmed double-booking authoritative save/reopen path represented by the final handoff commits

These areas are not implementation targets for the recovery phase.

## 4. Locked architectural contracts

### 4.1 Daily truth

Daily operational truth remains reconstructed from recurring/base schedule plus date-specific overlays and active replacement facts. Date-specific changes must not destructively rewrite the recurring programme.

### 4.2 User-facing architecture

- MASTER is the application home and a read-only operational hub.
- THERAPIST_DAILY and other views are projections, not business-rule owners.
- Internal storage and implementation details remain hidden from users.
- IDs remain internal even when names are shown in the UI.

### 4.3 Historical architecture

The historical model is explicitly hybrid:

- finalized daily snapshots for durable reporting
- append-only events for auditability and corrections
- explicit `Κλείσιμο ημέρας`
- finalization blocked while unresolved sessions or non-final outcomes exist
- corrections to closed days create amendment events and a new snapshot version
- previous snapshot versions remain preserved/superseded, never silently rewritten

### 4.4 Workload metric

The primary historical workload metric is the actual final operational workload for the day. Planned workload is a secondary comparison metric.

## 5. Immediate implementation milestone: statistics/history

The trusted handoff identifies this as the next milestone. The order is:

### Phase H1: Historical fact model and close-day validation

Build a pure-Python historical model first, with no workbook mutation and no new VBA.

Required behavior:

- represent a finalized day with explicit date and version
- preserve therapist/day actual operational workload
- preserve occupied timeslots and active-session count
- preserve final operational outcome counts
- reject finalization when unresolved/orphan sessions remain
- reject finalization when required sessions lack a final outcome
- support deterministic reconstruction from the authoritative current-day state
- define stable identity fields needed for later patient and therapist history

Acceptance gate:

- focused unit tests green
- no existing scheduling/replacement tests regressed
- no workbook file mutated

### Phase H2: Durable close-day persistence

Implement the explicit `Κλείσιμο ημέρας` persistence contract.

Required behavior:

- append finalized snapshot version 1 for a newly closed day
- preserve source/date/hash or equivalent integrity metadata needed for safe replay/audit
- idempotent/safe behavior when the same already-finalized state is encountered
- no silent overwrite of an existing closed-day version
- produce an auditable event describing finalization

Acceptance gate:

- pure-Python persistence tests green
- versioning tests green
- source workbook unchanged in preview/smoke preparation

### Phase H3: Correction of a closed day

Implement the dedicated closed-day correction path.

Required behavior:

- correction is explicit and separate from ordinary daily editing
- previous snapshot version remains intact
- amendment event records the correction
- a new snapshot version becomes current
- reports can distinguish current versus superseded versions

Acceptance gate:

- full version-chain tests
- no destructive rewrite of previous history

### Phase H4: First historical report slice

Implement therapist/day history first, as specified in the design checkpoint.

Required outputs:

- therapist
- day
- occupied timeslots
- active-session count
- final operational outcome counts
- actual workload
- planned workload where available as secondary comparison

Then add weekly therapist workload history and trend.

### Phase H5: Patient historical statistics

Required patient history includes, where the operational model supports it:

- scheduled sessions
- delivered sessions
- missed/no-show outcomes
- department cancellation / specific-session cancellation
- patient absence days and reasons
- replacement/suspension outcomes relevant to the patient
- later 15-day progress trend integration

### Phase H6: Therapist profile/statistics

Required therapist history includes:

- current workload count
- historical workload/trend
- absence days and reasons
- replacements / operational replacement history
- scheduled versus delivered workload measures where defined

### Phase H7: Dashboards, search and quick lists

Only after the historical fact model is stable:

- patient search/profile history
- therapist search/profile history
- quick lists
- statistics dashboard
- clickable therapist name from the app shell into therapist profile/statistics

Metric definitions must be locked before KPI presentation is finalized.

## 6. Patient progress milestone

After the historical layer is stable:

- 15-day main goal
- three progress states: progress / no significant change / negative progress
- append-only progress history
- patient trend/sparkline
- aggregate internal-patient progress view
- hospital-wide progress indicator only after its scoring contract is explicitly defined

Clickable treatment-time navigation from the patient/application view can later open this progress/evaluation surface.

## 7. Manual replacement UX requirement to preserve for later integration

The later user-approved manual operator contract remains a valid requirement, but it must be implemented on top of the proven replacement engine rather than by patching live VBA modules.

Required behavior:

- first visible outcome for each displaced patient: `Αναβολή θεραπείας για σήμερα`
- explicit provider choice
- explicit time choice from the currently valid times for that provider
- no silent first-time selection
- `Επιβεβαίωση επιλογής` confirms the current patient and does not auto-advance
- explicit `Προηγούμενος ασθενής` and `Επόμενος ασθενής`
- previously confirmed decision can be reviewed/revised
- revising an earlier decision re-evaluates downstream availability/capacity
- progress/status indicator showing confirmed/unconfirmed state
- every patient requires explicit confirmation
- `Αυτόματη τακτοποίηση υπολοίπων` remains available
- manually confirmed choices or date-specific suspensions remain fixed authoritative facts
- automatic solver acts only on the unresolved remainder
- if the remainder has no full solution, fixed manual decisions remain untouched and unresolved cases remain visible/manual
- manual preparation must stay lightweight and must not invoke a full global solve merely to display the form

This phase must reuse the protected candidate/policy/capacity engine. No forked replacement rules are allowed.

## 8. Future application requirements already recorded

The following remain in the backlog after the history foundation:

- patient name -> intranet/electronic folder using stable PatientID/FolderKey
- treatment time -> patient progress/evaluation view
- therapist name -> therapist profile/statistics dashboard
- shared multi-user intranet architecture with authoritative data, permissions and audit trail
- application-like Excel shell with hidden chrome where practical and safe
- modern guided forms
- room-grouped MASTER presentation and responsible doctor handling
- permanent responsible-doctor change workflow
- patient and therapist search
- quick lists
- patient-centric vertical workflow
- guided manual scheduling with therapist load ordering and only valid times exposed
- automatic scheduling/management using hard constraints plus scored preferences and explicit user confirmation before commit
- portable/exhibition build only after the operational application is stable
- final integrated beta after the major subsystems are complete

## 9. Replacement optimization requirements recorded in project decisions

When the project reaches the global replacement optimization milestone, preserve these rules:

- collect all relevant absences first
- solve displaced sessions together
- protect scarce feasible capacity
- prefer original time where feasible
- respect physiotherapy/reclining adjacency constraints
- balance workload
- use persistent pool/robotic operating hours
- weak preference to avoid 13:00 where possible
- explicit user-approved double booking only as a last-resort exception
- manual override always available

These are not to be reimplemented during the immediate history milestone unless a later trusted baseline proves that milestone had already been completed.

## 10. Workbook smoke policy

Native tests must use fresh outputs with monotonically new names, for example:

- `Excel/previews/RECOVERY_HISTORY_SMOKE_01.xlsm`
- `Excel/previews/RECOVERY_HISTORY_SMOKE_02.xlsm`
- `Excel/previews/RECOVERY_REPLACEMENT_SMOKE_01.xlsm`

Never reuse an old smoke as the mutation target for a new architectural experiment.

The full application smoke must be built through the canonical full pipeline, not by treating an intermediate preview such as `REGISTRATION_MENU_PREVIEW.xlsm` as the complete application.

## 11. Implementation sequencing from this recovery point

1. Freeze `recovery/known-good-4a2eed0`.
2. Create a child work branch for history/statistics.
3. Audit existing snapshot/statistics tests and reuse current authoritative reconstruction APIs.
4. Implement H1 as pure Python with tests only.
5. Run focused plus regression tests.
6. Implement persistence/versioning H2 and H3.
7. Only then add Excel integration for `Κλείσιμο ημέρας` on a fresh full-pipeline workbook copy.
8. Native proof on a newly named smoke workbook.
9. Build therapist/day history and weekly workload reports.
10. Continue through patient/therapist profiles and dashboards.
11. Revisit manual replacement UX only after the history milestone is stable, unless the user explicitly reprioritizes it.

## 12. Change-control rule

Any future conversation or handoff must start by reading this plan together with `CURRENT_HANDOFF.md`, `DECISIONS.md`, `PROJECT_STATUS.md`, `STATISTICS_HISTORY_DESIGN.md`, `ONBOARDING.md`, and `UI_APPLICATION_ROADMAP.md` before altering operational code.

If a new failure appears, identify the failing layer first. Do not use the failure as permission to redesign adjacent working systems.
