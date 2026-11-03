# Final hardening backlog: worksheet edit protection

Recorded from native user observation on 2026-11-03: users can currently type arbitrary text in many cells across user-facing worksheets. This is unacceptable for final production readiness, but is explicitly deferred until after the present H6 UX work to avoid destabilizing the proven save/history flows.

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

Priority: mandatory before final release, but **not now**.
