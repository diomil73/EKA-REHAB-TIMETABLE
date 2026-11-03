# H6 MASTER-only Save Day checkpoint | 2026-11-03

## Native proof
User confirmed AZ created through the H6 builder with the numeric Excel time comparator installer and unified Save Day installer. Using the MASTER Save Day button, first close succeeded. After another change and same-session save, user supplied:
```text
DAY AMENDMENT OUTCOME OK
Date: 2026-10-30
Version: 1 -> 2
Current snapshot: finalized:2026-10-30:v2
Therapist/day history rows: 16
Previous patient/session history rows: 80
Current patient/session history rows: 80
Operational CURRENT_FINAL patient rows: 80
```
Native v1 was reported successful; no separate v1 inspector output for AZ was pasted. H6 functional path is green with v2 inspector confirmation. **AZ GREEN/FROZEN**; do not modify.

## User UX feedback
Duplicate Save Day control in DAILY_INPUT is unnecessary, since Apply & Continue returns to MASTER and the user saves there. Screenshot also showed MASTER Save Day button overlapping Save & Exit.

## Scoped follow-up implemented, not yet native-tested
- `Python/tools/install_unified_day_save_ui.py` commit ca54a36: delete all DAILY_INPUT close/save shapes; create only MASTER Save Day after existing Save & Exit shape; verify no overlap, persisted bindings and no DAILY_INPUT duplicate.
- `Tests/test_unified_day_save_ui.py` commit 4332635: source regression guards.
- No changes to the authoritative close/amendment code or to the v1/v2 decision logic.
- Preserve production immutable and AX, AY, AZ green/frozen.
- Next validation: targeted pytest then fresh unique smoke and installation, read-only inspection of MASTER button placement; no need to redo the complete v1/v2 H6 functionality if the UI-only changes can be verified read-only. Mark this follow-up green only after native UI check.
