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

The lower daily-change area is a dedicated supervisory panel rather than a raw Excel table. It uses a restrained application-like presentation:

- pale mint panel background;
- stronger mint title/header zone;
- alternating very-light rows for readability;
- soft coral/pink restore actions;
- larger controls and increased row height;
- light borders instead of a dense spreadsheet grid.

The `Επιστροφή στο MASTER` control belongs in the same bottom action row as the restore controls. It must not sit above/right of the therapist schedule or float as a separate navigation element.

The bottom action row is calculated from the last visible daily-change row, so navigation and restore buttons remain grouped even when the ledger grows or shrinks.

The intended action-row order is:

`Επιστροφή στο MASTER | Επαναφορά επιλεγμένων | Επαναφορά όλων`

The presentation layer must be reapplied whenever the daily-change ledger refreshes; restore/business logic must not depend on the panel styling.

### Access policy

The dedicated therapist workload / management surface is intended only for:

- `SUPER_USER`;
- a future separately assigned `PHYSIOTHERAPY_HEAD` role.

Until authentication/role-aware navigation is implemented, the safe default is to provide no general MASTER entry point to this surface.

A future supervisor menu may call the existing `OpenTherapistDailyWithChanges` action after authorization succeeds.

## Daily changes / restore relationship

The reversible daily-change ledger remains part of the Physiotherapy supervisory workflow. Its navigation macro is retained, but it is no longer wired to a general MASTER workload button.

Navigation back to MASTER and patient restore controls are intentionally grouped into the same lower management zone.

This separation is intentional:

`MASTER visibility != therapist-management visibility`

## Implementation points

Relevant modules:

- `Python/rehab_excel/therapist_workload_strip_vba.py`
- `Python/rehab_excel/daily_change_ledger_vba.py`
- `Python/rehab_excel/daily_input_change_navigation_vba.py`

`daily_change_ledger_vba.py` remains authoritative for ledger/restore behavior. The dedicated presentation module `modSupervisorPanelStyle`, installed by `daily_input_change_navigation_vba.py`, repaints the legacy renderer after refresh and replaces grey restore Form Controls with styled Excel shapes. It also positions the MASTER navigation in the final action row. The installer injects the styling call into both the general ledger refresh path and the DAILY_INPUT synchronization path so the panel does not revert after refresh/restore.

Regression coverage:

- `Tests/test_therapist_navigation_and_locked_workload.py`
- `Tests/test_supervisor_panel_style_vba.py`

The production code must continue to remove legacy MASTER workload shapes from older previews/workbooks so an upgrade cannot accidentally leave management information visible.
