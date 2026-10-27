# Replacement operator reconnection checkpoint

Date: 2026-10-27

## Purpose

This checkpoint reconnects the already-established therapist-replacement subsystems without changing the protected replacement solver or business rules.

## Protected components

The following remain unchanged and authoritative:

- replacement feasibility and provider policy;
- patient cross-specialty conflict checks;
- provider and student capacity rules;
- service hours;
- reclining-treatment adjacency;
- sequential recalculation after a confirmed manual choice;
- deterministic global replacement planner and unique-slot protection;
- authoritative replacement facts and restore-by-ReplacementId lifecycle.

## Operator routing

When a therapist absence creates displaced sessions, the Excel operator must choose before placement is committed:

- **Yes**: calculate/use the existing global best solution for all unresolved sessions, with explicit confirmation before persistence;
- **No**: resolve patients one-by-one through the manual operator form;
- **Cancel**: stop without creating another replacement.

The DAILY_INPUT route is patched only by exact known markers and fails closed if the expected continuation point is not present.

## Manual path

The manual path is sequential:

1. prepare one unresolved patient using the existing feasibility/policy/service-hour pipeline;
2. show valid therapist/time choices;
3. require explicit operator confirmation;
4. revalidate the selected therapist/time against the current workbook state;
5. create one ACTIVE authoritative replacement fact;
6. rebuild operational projections;
7. synchronize the visible replacement ledger from authoritative facts;
8. recompute the unresolved queue before presenting the next patient.

The preparation and validation Python tools are read-only. They do not save or mutate the workbook.

## DAILY_CHANGE_LEDGER contract

`DAILY_CHANGE_LEDGER` remains the existing visible change/restore ledger beneath `THERAPIST_DAILY`. No second ledger is introduced.

The ledger is now treated as a projection of authoritative `DAILY_REPLACEMENT_FACTS`:

- `ChangeID` is the exact `ReplacementId`;
- patient, original therapist/time, effective therapist/time, reason and status are reconstructed from authoritative workbook data;
- ACTIVE and RESTORED history is preserved in the visible projection;
- the existing `RefreshDailyChangeLedgerView` renders `Από`, `Προς`, `Επαναφορά`, `Επαναφορά επιλεγμένων` and `Επαναφορά όλων`;
- the existing authoritative restore backend continues to restore by `ReplacementId` and full recomputation, not by cell snapshots.

Legacy snapshot columns remain part of the established sheet schema for compatibility/audit context, but they are not the source of truth for restore.

## Fail-closed wiring

The installer uses exact transforms to add ledger synchronization after:

- every confirmed manual replacement;
- completion of the selected replacement workflow, including the automatic-all route.

If expected VBA markers are absent or duplicated, the build must stop rather than apply a speculative patch.

## Native verification still required

No native Excel acceptance is claimed by this checkpoint. Before acceptance, run the regression gate, build a fresh preview from a trusted clean source, and verify this narrow sequence:

`therapist absence -> choose manual -> confirm one patient -> ledger row appears with From/To/Restore -> next patient options are recomputed -> restore the confirmed row`

Only after that manual smoke is green should the automatic-all route be smoke-tested.
