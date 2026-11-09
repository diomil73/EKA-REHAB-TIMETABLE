# Diomil | Dynamic therapist profile for 20+ therapists (design decision)

Status: **agreed functional direction; NOT IMPLEMENTED / NOT PILOT-APPROVED**.
User confirmed the therapist population may reach approximately 20 at some periods.
No separate worksheet per therapist and no duplication of VBA or formulas.

## User-facing contract
- Preserve current `Ιστορικό / Στατιστικά Θεραπευτών` as department/team comparison dashboard.
- Selection of a therapist row automatically updates the existing charts; do **not** require a separate `Γράφημα θεραπευτή` button.
- A future **single dynamic `Καρτέλα θεραπευτή`** view will show the selected therapist's individual workload, treatments, longitudinal trends, working/absence calendar and operational replacement history, subject to authorization.
- Retain `Στατιστικά θεραπευτών` on management screen as the navigation link back to the team statistics. This is *not* the redundant per-person chart button.
- Support 20+ therapists by data-driven registry and list selection, NOT by 20 sheets. No fixed list cardinality of six or twenty.
- Active, temporarily unavailable, and inactive status must not erase historical data or change immutable therapist IDs; daily assignment eligibility is distinct from membership in registry.

## Data-model direction, not implementation authority
- Authoritative therapist registry: immutable therapist ID, display name, operational status/effective dates. Avoid joining facts on mutable labels.
- Separate append-only absence ledger (therapist ID, date/range, broad operational category, status, source, provenance and amendments). Avoid multiple records from separate screens for the same absence; enforce stable event identity and audit trail.
- Day scheduler should consume verified *operational availability*, not infer medical reason or automatic removal of planned sessions. Replacements/cancellations retain independent patient-session audit events.
- Attendance % and working-day denominators remain **unavailable until independently verified**; synthetic H8 attendance is not actual clinical/payroll attendance.
- Medical/health-related reasons are sensitive; the common dashboard/management export should expose neither diagnostic detail nor free-text private explanation. Profile and writable attendance forms require separate roles/access control, minimized category visibility and local storage/privacy review before any implementation.
- Export/reporting uses aggregates with patient/therapist access restrictions as appropriate.
- Do not add attendance or personal detail to frozen H10 demo BZ, and do not introduce a production absence module in current testing cycle.

## Current H10 operational-first cleanup (source-only)
- Remove redundant top toolbar `Γράφημα θεραπευτή` (row auto-selection already changes charts). Preserve `Στατιστικά θεραπευτών` on management view as navigation.
- Expand workload sparkline vertical height while keeping 0–6 / step 1; no fabricated source values, no chart overlap acceptance without a new Excel screenshot.
- Separate management KPI cards with distinct gutters and centered content. A later visual acceptance remains necessary.
- Installed BZ is frozen test evidence and DOES NOT include these source-only changes. A fresh uniquely named future candidate is needed after source tests and installer gate adjustment, not an overwrite.
