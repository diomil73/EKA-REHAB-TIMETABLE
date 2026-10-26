# Therapist Workload Privacy and Navigation UI

Status: agreed and implemented on branch.

## Purpose

Therapist workload is management information. It must not be exposed merely because a user is allowed to view `MASTER_SCHEDULE`.

The general MASTER surface is shared by users with very different responsibilities, including medical, nursing and transport staff. Access to MASTER therefore never implies permission to inspect therapist workload.

## Locked UI contract

### MASTER_SCHEDULE

`MASTER_SCHEDULE` must not contain:

- therapist workload badges;
- therapist workload counts;
- a `Φόρτος θεραπευτών` button;
- a general navigation path that reveals the therapist management view.

Legacy `ekaWorkload_*` shapes are removed during installation/migration and again defensively by the workload renderer.

### THERAPIST_DAILY

Therapist workload remains available on the dedicated supervisory surface `THERAPIST_DAILY`.

The workload badges are rendered beside the operational grid, beginning at column K, and are not rendered on MASTER.

The `Επιστροφή στο MASTER` navigation control is anchored at `K1`, beside the top of the operational grid. It must not be placed in the middle separator area of the schedule.

### Access policy

The dedicated therapist workload / management surface is intended only for:

- `SUPER_USER`;
- a future separately assigned `PHYSIOTHERAPY_HEAD` role.

Until authentication/role-aware navigation is implemented, the safe default is to provide no general MASTER entry point to this surface.

A future supervisor menu may call the existing `OpenTherapistDailyWithChanges` action after authorization succeeds.

## Daily changes / restore relationship

The reversible daily-change ledger remains part of the Physiotherapy supervisory workflow. Its navigation macro is retained, but it is no longer wired to a general MASTER workload button.

This separation is intentional:

`MASTER visibility != therapist-management visibility`

## Implementation points

Relevant modules:

- `Python/rehab_excel/therapist_workload_strip_vba.py`
- `Python/rehab_excel/daily_input_change_navigation_vba.py`

Regression coverage:

- `Tests/test_therapist_navigation_and_locked_workload.py`

The production code must continue to remove legacy MASTER workload shapes from older previews/workbooks so an upgrade cannot accidentally leave management information visible.
