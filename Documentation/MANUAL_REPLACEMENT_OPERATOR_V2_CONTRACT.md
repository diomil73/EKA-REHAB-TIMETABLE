# Manual replacement operator V2 contract

Date: 2026-10-29

## Scope

This document locks the operator-facing contract for therapist-absence manual replacement handling. It does not replace the protected feasibility, capacity, service-hours or global-solver rules.

## Core principle

The operator controls when the workflow moves from one displaced patient to another. A confirmation must never silently advance the UI to the next patient.

## Required patient outcomes

For every displaced patient, the first operational outcome shown must be:

- **Αναβολή θεραπείας για σήμερα**

This represents a date-specific treatment suspension for a patient who will not be placed that day. It is not a permanent cancellation and it must participate in the same authoritative daily-state/recompute lifecycle as other daily operational facts.

The normal alternative is a replacement placement consisting of both:

- an explicitly selected replacement therapist/provider; and
- an explicitly selected replacement time from that provider's currently valid available times.

The UI must never silently choose the first available time merely because a provider was selected.

## Explicit therapist/time selection

The manual form must present therapist/provider selection separately from time selection.

When a provider is selected, only that provider's currently valid available times are exposed. The operator must explicitly select one of those times before confirmation.

Capacity/load information remains visible as decision support, but it must not substitute for the explicit time choice.

## Confirmation and navigation

`Επιβεβαίωση επιλογής` records/revalidates the current patient's decision but does not change patient automatically.

After confirmation the form remains on the same patient and clearly shows that the current decision is confirmed. The operator may correct the decision before choosing to navigate away.

The form must provide explicit navigation controls:

- **Προηγούμενος ασθενής**
- **Επόμενος ασθενής**

Moving to another patient is therefore an operator action, not a side effect of confirmation.

Any backward edit that invalidates downstream capacity must trigger authoritative recomputation of affected unresolved choices. Previously confirmed manual facts remain authoritative until explicitly changed/restored.

## Automatic completion of the remainder

At every point in manual mode the form must expose:

- **Αυτόματη τακτοποίηση υπολοίπων**

This action runs the existing protected global replacement solver only for the still-unresolved patients. Manual decisions already confirmed by the operator, plus date-specific suspensions already confirmed, are treated as fixed daily facts and must not be silently reassigned.

This permits a mixed workflow: the operator can place selected special cases manually and then delegate the remaining unresolved patients to the global solver.

## Progress/status feedback

The form must make workflow state visible, including at least:

- current patient position or equivalent progress indicator;
- number of unresolved sessions/patients;
- whether the current decision is unconfirmed or confirmed;
- clear status while a recomputation is running.

The UI must not leave the operator guessing whether Excel is working, frozen, or has advanced to another patient.

## Protected business rules

V2 must reuse, not fork, the established rules for:

- provider eligibility;
- therapist/student capacity;
- patient cross-specialty conflicts;
- service hours;
- current ACTIVE daily replacement facts;
- current daily absences/cancellations/suspensions;
- authoritative global solve for automatic remainder placement.

Manual preparation remains lightweight and must not run the full global solver merely to open or decorate the current patient's form.

## Runtime protection finding from native smoke

After the manual form became responsive, confirming a choice reached the visible ledger refresh and Excel reported that the cell/chart being changed was on a protected sheet.

The concrete renderer involved is `RefreshDailyChangeLedgerView` on `THERAPIST_DAILY`. The renderer writes cells and creates FormControls, so it must own a narrow live-session unprotect/render/reprotect bracket. Excel does not persist `UserInterfaceOnly=True` across close/reopen sessions.

The protection repair is intentionally isolated from replacement feasibility and solver logic. Native acceptance is required before the V2 UI is installed into the smoke workbook.
