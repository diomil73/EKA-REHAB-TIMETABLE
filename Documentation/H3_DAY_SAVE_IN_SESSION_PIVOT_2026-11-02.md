# H3 Day Save In-Session Pivot — 2026-11-02

## User-facing contract

`Αποθήκευση ημέρας` must complete silently inside the already-open Excel application.

It must not:
- close the workbook,
- quit Excel,
- reopen the workbook,
- ask for a manual reason code,
- ask for a manual summary.

The user confirms only that the closed day should be saved with its current changes.

## Historical evidence

AF failed because post-save verification saw no persisted amendment after a read-only COM fallback.
AG failed because the workbook remained read-only/locked during worker retry.
AH exposed a Windows `CreateFile` sharing violation in the release probe.
AI proved the workbook remained locked for 60 file-lock probes.
AJ finally proved the close handoff itself could work, but apply was rejected because Excel rewrote the `.xlsm` package during close, making the pre-close payload hash stale.

AJ exact worker outcome supplied by operator:

```text
{
  "ok": false,
  "error": "Workbook changed after amendment preparation; refusing to apply payload"
}
```

AJ inspector outcome supplied by operator:

```text
DAY AMENDMENT INSPECTION FAILED: Expected exactly one amended service date, found 0
```

AJ is frozen and must not be reused.

## Architecture decision

The close/reopen worker architecture is retired from the user-facing H3 flow.

The day-save flow now mirrors the already-proven close-day architecture:

1. User confirms `Αποθήκευση ημέρας`.
2. The open workbook is saved.
3. Python performs read-only calculation/validation and emits a TSV payload.
4. VBA persists the complete version transition inside the same open workbook.
5. VBA performs read-back verification.
6. On any persistence failure, VBA rolls back the in-session mutation.
7. The workbook is saved again.
8. Excel remains open throughout.

The transition remains versioned and append-aware:
- prior snapshot becomes `SUPERSEDED`,
- replacement snapshot becomes new `CURRENT_FINAL`,
- therapist/day history rows are written for the replacement snapshot,
- linked STARTED and FINALIZED history events are written automatically,
- no user-authored reason/summary is required.

## Source changes

- `43d6fdad960bf54f3e775dace6c38b5d1ceb33d1` — add in-session TSV preparation helper
- `492bbfac597e3a8574b8b4b363d47781a6797e58` — replace close/reopen H3 action with in-session VBA persistence
- `7c7d541f9b67ad73c9ac08311eab187c01eae8c9` — assert no close/quit/worker/reopen in user-facing action
- `5abfd2e1ec7b197eafcb96e508dc981270b32648` — cover in-session payload contract
- `95a9debe58de9dae94a2a06a4b7979686f049e6b` — make helper directly executable from VBA/PowerShell

Earlier worker commits remain in history but are no longer part of the user-facing H3 path.

## Next proof

Run the focused source gate. If green, create a fresh native workbook from frozen Z. Do not reuse AF/AG/AH/AI/AJ.
