# DIOMIL — Afternoon physiotherapy, pool and medical deferral scope (2026-11-15)

Status: **pure policy and synthetic tests committed on an isolated branch; NOT operationally integrated**.

## Operator-confirmed requirements

1. The 2026-11-12 real workbook is only a source of refreshed patient/treatment facts. Never replicate its layout inside the canonical DIOMIL MASTER. The existing doctor-under-patient display convention remains unchanged.
2. Source column map for the legacy `ΠΡΟΓΡΑΜΜΑ` sheet: A clinic/ward/doctor, B patient, C/D/E morning physiotherapy time/days/provider, F tilt-table time, G/H/I occupational therapy time/day/provider, J/K/L speech therapy, M/N/O psychological therapy, P/Q EFA time/days, R/S/T robotics time/days/provider, U afternoon programme (do not import), V pool time. The actual source positions are unchanged; any future DIOMIL UI layout is independent.
3. Pool is a **morning** treatment. The existing `Πισίνα` weekly treatment code already exists in `trial_admission_gate.py` and `weekly_slots_source.py`. Do not duplicate this taxonomy or guess missing day/provider.
4. Afternoon programme is **physiotherapy only**. Its patient-management fields are `time`, `day`, `physiotherapist`. They start **blank**; do not import unreviewed legacy column U or automatically create any PM session.
5. Morning and afternoon schedules use **independent conflict namespaces**. A shared time, patient or therapist across the two programmes does not count as a cross-programme scheduling collision. Within each programme, reject overlapping sessions sharing a patient or assigned therapist, subject to the existing explicit group-therapy approval contract when integrated.
6. Temporary discharge and transfer are patient-wide states covering both morning and afternoon, for the relevant effective dates. This initial pure policy models **one explicit date** only, not a date range.
7. Medical deferral requires an explicit scope:
   - whole day: all scheduled patient therapies including afternoon;
   - specific time window: only sessions overlapping that actual window (and optionally a selected programme);
   - specific session: exactly the selected session ID, not every physiotherapy session.
   Do not infer duration from start times. Never erase recurrence rows because of a one-day deferral.
8. Morning-specific physiotherapy deferral, or absence of the morning physiotherapist, does not imply cancelling afternoon physiotherapy with its separately assigned therapist. The converse is also true.
9. A user-confirmed cancellation must be previewed with the exact affected sessions before save. Record reason, scope, date, targets, time bounds and restore/amendment history. No silent retrospective rewrite.

## Implementation slice committed

- `Python/rehab_core/programme_scope_policy.py`: pure immutable scheduling and absence records; `affected_session_ids` preview and `conflicting_session_pairs` local namespace checker.
- `Tests/test_programme_scope_policy.py`: **synthetic-only** regression scenarios; no patient names or identifiers from real data.
- Both files do **not** modify any Excel workbook, VBA, MASTER, patient registry, CC or frozen native acceptance artifact.
- This is a policy demonstration, **not** installed in DAILY_INPUT, UI, patient management, scheduling persistence or stats. Do not claim end-to-end completion.

## Integration tasks before operational release

1. Read the existing patient manager forms and insertion/edit macro source. Add optional afternoon-physio controls and morning-pool option only where missing, without changing the visual MASTER layout or treatment legacy semantics.
2. Keep PM fields blank for every imported patient; require explicit date/day, clock and therapist before saving an afternoon appointment. Validate the three fields together; do not infer absent values from AM physiotherapy.
3. Bridge operator selections to an established daily-change ledger and reversible day history, after confirming current protected-sheet write paths. Do not overwrite or bypass existing daily-close/version-chain guarantees.
4. Reconcile morning and afternoon duration semantics and group/shared therapist sessions from existing application rules before reusing the pure conflict function for real data.
5. Add two-stage preview and confirmation of affected sessions for full-day or medical window deferral, then exact native Windows Excel smoke on a **fresh uniquely named disposable XLSM**, with independent read-only outcome inspection.
6. Include afternoon-specific statistics without contaminating morning workload, and ensure cancellations record accurate treatment/time/provider outcomes. Keep current daily/weekly/historical data unchanged.
7. All real source staging, identity matching and patient data remain **local-only and private**; repository is public. No raw source, hashes used as source fingerprint if confidential, names, row dumps or patient details in Git/CI.
8. Preserve: production baseline, original 2026-11-12 source, frozen synthetic CC and every accepted/failed/uncertain smoke XLSM. No overwrite, repair-in-place or `git clean`.
9. Perform focused synthetic regression suite, full repository suite, then actual Excel native user acceptance. Tests are not considered green without output evidence.

## Focused local test command (repository root)

```powershell
git fetch origin
git switch work/afternoon-medical-scope-20261115
$env:PYTHONPATH = (Join-Path (Get-Location).Path "Python")
python -m pytest Tests/test_programme_scope_policy.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: afternoon/medical policy tests failed" }
```

Before switching, run `git status --short` and ensure tracked work is preserved. Untracked local previews/forensic backups must remain untouched. For source state the previous main work branch remains `work/history-statistics-rebuild`. This feature branch must not be merged until coverage and native integration gates are satisfied.

## Troubleshooting / stop conditions

- Missing or ambiguous therapist, time/day or programme: do not create afternoon record.
- Afternoon specialty other than physiotherapy: refuse.
- Patient state full-day vs targeted session: never infer based on free-text reason; choose scope explicitly.
- Target absent or session id belongs to another patient/date/programme: refuse.
- A patient has no afternoon entry: no phantom afternoon cancellation; full-day still cancels valid morning entries.
- Two sessions share a time across AM and PM: **not** a conflict based on cross-programme overlap alone.
- Two sessions overlap within the same programme: apply the existing clinical/shared-session approval policy before operational rejection or admission.
- Historical day is finalized: use amendment/version chain, not direct edit.
- Any native save, restore or read-back failure: stop once, preserve evidence; no second attempt on the same smoke copy.
