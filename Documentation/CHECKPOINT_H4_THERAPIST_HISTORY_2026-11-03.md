# H4 Therapist History / Statistics Checkpoint

Date: 2026-11-03
Branch: `work/history-statistics-rebuild`

## Status

H4 is operator-confirmed GREEN and frozen.

Canonical green native UI artifact:

`Excel/previews/UNIFIED_THERAPIST_HISTORY_UI_20261102_AQ.xlsm`

AQ must not be repaired or reused for future native smoke work.

## Proven reporting backend

The historical therapist reporting layer reads only authoritative finalized history and counts only the `CURRENT_FINAL` snapshot for each service date. Superseded historical versions remain preserved for audit but are not double-counted in reports.

Read-only report views implemented and previously operator-confirmed on the green AM history artifact:

- daily,
- weekly,
- monthly,
- overall.

The report facts include:

- therapist,
- service day / period,
- planned occupied timeslots,
- actual occupied timeslots,
- planned sessions,
- active sessions,
- replacements received,
- released load,
- treatment suspensions affecting load,
- planned utilization,
- actual utilization.

## Proven Excel UI

User-facing sheet:

`THERAPIST_HISTORY_REPORT`

Entry point:

`THERAPIST_DAILY` -> `Στατιστικά θεραπευτών`

The report stays inside the same open Excel session and does not expose the technical history sheets.

User-facing modes:

- `Ημερήσια`
- `Εβδομαδιαία`
- `Μηνιαία`
- `Συνολικά`
- `← MASTER`

Technical snapshot version is intentionally hidden from the user-facing report.

## Native UI evidence

Earlier candidates AN, AO and AP are failed/frozen UI evidence.

AQ is the first final green UI proof after the button-persistence fixes.

Operator-confirmed sequence in AQ:

1. Ημερήσια
2. Εβδομαδιαία
3. Μηνιαία
4. Συνολικά
5. Ημερήσια

Observed result: all views worked and the complete button row remained visible and usable through the mode switches.

## Final UI hardening commits

- `591ddbb` — reset therapist-history window state across report modes
- `4deb414` — regression coverage for the window reset / button visibility behavior

## Safety / continuity

- AQ is GREEN / FROZEN.
- H4 should not be reopened unless new concrete evidence identifies a specific defect.
- Technical history sheets remain hidden/protected.
- Reporting remains read-only over durable finalized history.
- Do not count superseded day snapshots in operational historical totals.

## Next roadmap slice

H5 — patient historical statistics:

- scheduled vs delivered,
- missed / no-show,
- department or specific-session cancellation,
- patient absence days and reasons,
- replacement / suspension outcomes,
- later 15-day progress trend.
