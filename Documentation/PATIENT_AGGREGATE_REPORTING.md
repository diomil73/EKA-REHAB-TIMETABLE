# Patient Aggregate Reporting Scope

Last updated: 2026-10-24

## Purpose

This document locks the patient-side aggregate reporting requirement that complements the individual patient history/progress model already defined in `Documentation/STATISTICS_HISTORY_DESIGN.md`.

## Reporting levels

Patient reporting must support three linked levels built from the same finalized historical facts:

1. **Individual patient**
   - scheduled and performed treatments,
   - missed/lost treatments with structured reasons,
   - treatment days and treatment types,
   - replacements/suspensions that affected the patient,
   - stay duration and treatment timeline,
   - 15-day progress/evaluation milestones with `Από`, `Έως` and green/grey/red outcome,
   - chronological progress trend.

2. **All patients / whole rehabilitation department**
   - number of patients served in the selected period,
   - treatment/session totals,
   - scheduled vs performed,
   - missed/lost treatments and reason breakdown,
   - replacements and treatment suspensions affecting patients,
   - treatment-type mix,
   - patient progress distributions/trends only where the aggregation has a clinically meaningful rule.

3. **Clinic-level patient aggregates**
   - the same aggregate metrics must be filterable/groupable by the patient's clinic/origin clinical unit,
   - clinic totals must reconcile exactly with the all-patient total for the same period and filters,
   - users must be able to compare clinics side-by-side and drill from a clinic total to the contributing patients and finalized historical records.

## Data model rule

Clinic is a reporting dimension attached to stable patient identity/care context, not inferred from today's worksheet position or visible text.

The implementation must use a stable clinic identifier/code internally and a normal business-facing clinic name in the UI. If clinic assignment can change during a patient's stay, historical reporting must preserve the clinic that was valid for the relevant care period/event rather than retroactively rewriting past reports with the patient's current clinic.

## Progress aggregation safety

Individual 15-day progress remains the authoritative clinical detail.

Department/clinic progress summaries must **not** use a naive arithmetic average of green/grey/red states. Before an aggregate progress indicator is implemented, its clinical meaning and calculation rule must be explicitly defined and documented. Until then, safe aggregate views may show counts/distributions of milestone outcomes and trends by period.

## Reporting consistency

The patient reporting hierarchy is therefore:

`Individual patient -> Clinic aggregate -> All-patient department aggregate`

All three levels must use the same finalized snapshots/events so totals always reconcile. No separate manually maintained statistical dataset is allowed.

## Scope note

This patient-side aggregate reporting is separate from therapist workload reporting, but both belong to the same statistics/history subsystem and must share the same finalization/versioning model.
