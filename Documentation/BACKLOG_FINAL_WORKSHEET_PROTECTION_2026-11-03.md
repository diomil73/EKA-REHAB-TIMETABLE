# Final hardening backlog: worksheet edit protection

Recorded from native user observation on 2026-11-03: users can currently type arbitrary text in many cells across user-facing worksheets. This is unacceptable for final production readiness. The H7 history/statistics UX and native smoke are now user-confirmed functional, so this work is ACTIVE as the next hardening milestone. Start with a read-only inventory on the accepted disposable BF; do not change protection until intended edit zones are explicitly mapped and regression contracts are in place.

## Acceptance and scope

- MASTER_SCHEDULE: read-only data surface, with working navigation and contextual Save Day buttons. No arbitrary cell edits.
- DAILY_INPUT (Absences/Cancellations): only intended data entry fields and approved form/workflow edits available. Keep Apply & Continue, save, protection refresh, validation and rollback functioning.
- THERAPIST_DAILY and other operational views: shield calculated/projected data from manual writes; preserve intended user actions.
- Registry and other underlying records: edits only through sanctioned forms and authoritative macros.
- FINALIZED_DAY_SNAPSHOTS, THERAPIST_DAY_HISTORY, PATIENT_SESSION_HISTORY and DAY_HISTORY_EVENTS: retain veryHidden and protection, with no direct user editing and stable authoritative VBA writes/readbacks.
- Review Excel sheet protection, locked/unlocked cell ranges, selection permissions, workbook structure visibility, button interactions and macro writes.
- Re-protect correctly on open, after session changes, save, and rollback/error; avoid Excel UserInterfaceOnly persistence assumptions and any broad unprotect window.
- Regression and native acceptance must prove attempted arbitrary cell edit is blocked while intended entry, first close, amendment v1→v2, and same-session workflows remain intact.
- Never change production workbook or already green/frozen smoke artifacts. Changes require focused tests and fresh native smoke.

Priority: mandatory before final release. **Current status: ACTIVE, read-only baseline audit first.**


## 2026-11-04 status update

H7 is now source-green and user-confirmed in native BF:
- 26 focused tests passed.
- BF integrated read-only audit passed.
- User manually confirmed the toolbar, RefreshAll, patient/therapist reports and previous-screen navigation all worked correctly.
- BF remains disposable smoke evidence, not production.

Next work is the protection hardening above. First diagnostic:
`Python/tools/audit_workbook_protection.py` (read-only; commit `aa0cd3c`, permission-report enhancement `cf04e93`) inventories workbook structure protection, per-sheet content/drawing/scenario protection, selection permissions, UsedRange lock state, and unlocked cell ranges on protected sheets. It never calls Protect, Unprotect or Save. It opens the supplied workbook read-only and refuses incomplete large-range scans rather than claiming full coverage.

First target: accepted BF only:
`Excel/previews/UNIFIED_PATIENT_HISTORY_H7_20261104_BF.xlsm`.

Do not implement blanket protection based only on this audit. Use results to classify each sheet and map allowed edit zones for DAILY_INPUT and sanctioned forms/macros. Preserve all native write/readback, close-day v1, amendment v1→v2, refresh and history/report flows. Every protection change must be tested on a new uniquely named workbook copied from the trusted frozen source; never patch BF or production.
