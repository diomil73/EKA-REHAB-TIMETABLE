# H7 BD registry coverage reconciled | 2026-11-04

## Actual operator output (read-only BD)

```text
H7 REGISTRY VS CURRENT_FINAL READ-ONLY AUDIT
PATIENTS registry entries: 93
CURRENT_FINAL sessions: 80
Distinct patient IDs in CURRENT_FINAL: 77
Report rows: daily=77, weekly=77
ID 15: registry=YES finalized_sessions=NO daily=NO weekly=NO
ID 18: registry=YES finalized_sessions=YES daily=YES weekly=YES
ID 19: registry=YES finalized_sessions=NO daily=NO weekly=NO
ID 20: registry=YES finalized_sessions=NO daily=NO weekly=NO
ID 21: registry=YES finalized_sessions=NO daily=NO weekly=NO
ID 22: registry=YES finalized_sessions=YES daily=YES weekly=YES
Registry numeric IDs without CURRENT_FINAL sessions: 02, 06, 15, 19, 20, 21, 24, 25, 30, 39, 41, 50, 64, 69, 73, 75
CURRENT_FINAL numeric IDs absent from registry: none
Workbook unchanged: True
H7 REGISTRY AUDIT COMPLETE
```

### Interpretation and implications

The existing H7 report is a finalized-activity report, not a complete patient-registry list. All 80 sessions reconcile across all report modes, no historical records were lost. The absent 16 are present in the registry but have no CURRENT_FINAL session on the one tested date. The 18 and 22 are present in the report data; their apparent omission in Excel needs a separate **nonmutating worksheet row visibility/filter/scroll audit**.

A future option to show all registered patients may display explicit zero-activity rows. This must be distinctly identified as a *current registry view*, not an immutable historical roster; do not create fake finalized sessions or change 80/77/3 counts. Preserve existing activity-only report and exporter APIs as defaults, and add an explicit opt-in full-roster view.

### Pending native defects

- BD integrated persisted-structure audit GREEN, but intermittent generic failure on Overall was operator-reported before subsequent successful clicks. Root cause unconfirmed. Staged unmerge-before-clear plus error detail source fix still requires fresh native candidate.
- Paired therapist/patient stats buttons misaligned with one white-text label; code correction source-gated 6 passed in 1.78s, but not yet native-installed/verified.
- BD diagnostic/FROZEN, no more installer runs on BD. AZ and BB GREEN/FROZEN.
- Mandatory final worksheet edit protection gate before release.
