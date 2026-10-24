# Student presentation and lifecycle contract

Last updated: 2026-10-24

This document records the locked user-facing presentation and lifecycle behavior for physiotherapy students. It complements the authoritative `STUDENTS` registry and the unified therapist/student replacement engine.

## Active placement presentation

While a student placement is active:

- the student is shown to users by the student's real display name, not by a generic slot label;
- the system must still preserve `ReplacementProviderKind.STUDENT` and the stable StudentID internally, so students are never treated as therapists just because they appear in the same operational views;
- `THERAPIST_DAILY` should show the student's real name using a distinct student-specific visual treatment/color so staff can immediately recognize that the provider is a student;
- workload/capacity views must include the student separately and must also identify the provider visually as a student rather than presenting the name exactly like a therapist;
- the student-specific color is presentation metadata only. Scheduling, workload and replacement rules remain driven by authoritative provider kind and StudentID.

## THERAPIST_DAILY physical slots

The current workbook contains numbered physical student slots such as `φοιτ 1`, with room for future numbered slots (`φοιτ 2`, `φοιτ 3`, ...).

These numbered labels are physical/grid placeholders, not the normal active-placement display name. During an active placement, the assigned slot should present the student's real name with student styling while the system retains the deterministic mapping to the numbered physical slot.

If active students exceed available physical student slots, the system must safety-stop rather than merge multiple students into one slot or silently treat them as therapists.

## End of placement

When a student's placement ends:

- the student's real name must no longer remain as the active label in `THERAPIST_DAILY` or workload presentation;
- the physical slot returns to its neutral generic placeholder such as `φοιτ 1`, `φοιτ 2`, etc.;
- historical facts must continue to preserve the real StudentID/display identity for audit/history and must not be rewritten to the generic placeholder;
- any recurring/session ownership that still points to the expired student must be resolved through the existing student-end transition/reassignment workflow rather than hidden by the label change;
- the generic placeholder indicates an available student position, not a provider identity.

## Workload semantics

Students participate in operational workload/capacity while active, using the student-specific capacity rules already defined by the scheduling engine. Their workload presentation should make the student status visible through the same student-specific visual language used in `THERAPIST_DAILY`.

A future UI pass may choose the exact color, typography or badge treatment, but it must preserve the semantic distinction between therapist and student and should use one consistent student presentation across operational views.

## Production architecture rule

There is no separate student-replacement production system. Students remain candidates inside the common global replacement engine. The student-specific behavior is expressed through provider kind, registry/lifecycle rules, capability/capacity rules and presentation mapping.

Smoke/preparation/verifier scripts dedicated to students are development/test tooling only and must not become a parallel production workflow.
