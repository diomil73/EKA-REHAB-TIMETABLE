# Recovery Replacement Operator Restore — 2026-10-31

## Native G findings that block close-day

`UNIFIED_CLOSE_DAY_RECOVERY_20261031_G.xlsm` remains frozen evidence and must not be rebuilt, overwritten or repaired in place.

Native G exposed three contract regressions:

1. Patient absence removed the MASTER treatment text instead of retaining the recurring programme with the agreed red strike-through presentation. The patient status also remained `παρών`.
2. Therapist absence did not expose its recorded reason in the replacement decision UI.
3. The recovered global replacement action showed only a Yes/No prompt. Selecting No did not open the previously established guided manual workflow.

`CloseOperationalDay` remains blocked until a fresh artifact verifies the restored contracts.

## MASTER daily presentation restored from established feature source

The established implementation on `feature/outpatient-scheduling` was recovered rather than redesigned.

Commits:

- `90ee9bc` restores the established MASTER daily non-active overlay.
- `51b31fc` restores the corresponding regressions.

Locked behavior:

- `PATIENT_ABSENT`, `CANCELLED`, and unresolved `THERAPIST_ABSENT` keep the recurring treatment text in place.
- The affected MASTER treatment cell receives a `PRESENTATION` patch with strike-through and source tag `master_daily_non_active_overlay`.
- Native DAILY_INPUT apply renders that source tag as red strike-through.
- The `Κατάσταση` column is discovered dynamically and receives the dated reason/fallback status when no active/replaced session remains for that inpatient on the selected date.
- Recurring/base schedule rows are never edited.
- Accepted replacement continues to show effective provider/time.

## Stable replacement operator foundation recovered

The stable `modReplacementOperator` / `frmReplacementDecision` design from `feature/outpatient-scheduling` is the recovery foundation. Quarantined V2 patch/normalizer/repair tooling is not used.

Read-only support recovery:

- `05a34e8` restores replacement decision preparation and exposes the therapist-absence reason.
- `16c9bcb` restores read-only revalidation of a chosen provider/time against the same provider policy, service hours, availability and capacity rules.
- `95d2673` adds deferred-target support for explicit `Επόμενος` navigation without marking that patient resolved.

Native operator recovery:

- `d162ec4` restores `modReplacementOperator` and `frmReplacementDecision` as an in-session operator layer.
- top-level choice is again `Ναι = καλύτερη αυτόματη συνολική λύση`, `Όχι = χειροκίνητα ασθενή-ασθενή`, `Άκυρο = διακοπή`;
- per-patient actions are explicit replacement, `Αναβολή για σήμερα`, and `Επόμενος` without completion;
- provider and time are explicitly selected from validated options;
- therapist-absence reason is shown in the form;
- `Προηγούμενος` undoes the latest manual replacement/suspension and refreshes the operational projection before recalculation;
- `Αυτόματη τακτοποίηση υπολοίπων` may be called from manual mode at any time and calls the existing global solver while already persisted manual decisions remain fixed facts;
- manual daily postponement uses the established DAILY_INPUT `Αναστολή θεραπείας` semantics and is day-specific;
- operator projection refresh uses the established in-session prepare/apply/readback path and does not render workload UI on MASTER;
- no `ThisWorkbook.Close`, `Application.OnTime`, authoritative reopen or V2 tooling is introduced.

## Global automatic action ownership

`abe82b5` wires the operator through the existing operational build patch:

- DAILY_INPUT is routed from direct `ResolveDailyReplacementsInPlace()` to `ResolveReplacementWorkflow()`;
- global automatic action remains the authoritative solver/apply path;
- its internal Yes/No confirmation no longer falsely claims that No starts manual mode;
- current-action replacement summary remains `patient -> therapist -> time`;
- workload remains outside the MASTER surface.

## Application contract

`156fdd7` promotes the operator module and form into the canonical application contract.

Required VBA component count is now **23**, adding:

- `modReplacementOperator`
- `frmReplacementDecision`

`7995804` updates application-contract regressions.

## Recovery regressions

- `78c6cbd` adds focused operator recovery regressions for top-level auto/manual/cancel routing, three per-patient outcomes, reason visibility, previous/next behavior, automatic remainder, read-only support tools, and in-session/no-workload refresh.
- `a844601` aligns visibility regressions with operator ownership and verifies that the global automatic prompt no longer impersonates the manual route.

## Solver boundary

No global solver scoring or feasibility rule was changed in this recovery slice.

The existing authoritative rules remain in force, including:

- combined/global optimization of all unresolved displaced sessions,
- hard availability/conflict/service/capacity constraints,
- scarce-slot protection,
- original-time preference,
- physiotherapy/reclining adjacency preference,
- workload balance,
- weak 13:00 avoidance,
- explicit review for double-booking/over-capacity exceptions.

## Next gate

Before any H workbook is built, run the focused source gate covering:

- MASTER daily overlay,
- DAILY_INPUT in-place preparation/action,
- replacement operator source and support tools,
- operator visibility/global replacement integration,
- production in-session safety,
- application contract,
- unified builder regressions.

If green, build a fresh uniquely named H artifact. H native testing must verify patient-absence red strike/status reason and the restored replacement operator before `CloseOperationalDay` is allowed.
