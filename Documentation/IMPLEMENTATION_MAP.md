# Implementation Map

Τελευταία ενημέρωση: 2026-10-20

## Σκοπός

Αυτό είναι ο ζωντανός χάρτης ownership της υλοποίησης. Για συνολική κατανόηση του project ξεκινάμε από `Documentation/ONBOARDING.md`. Για το ακριβές resume point χρησιμοποιούμε `Documentation/CURRENT_HANDOFF.md`.

## Κύρια ownership περιοχών

### Scheduling / business logic

- `Python/rehab_core/availability.py`: patient/provider availability.
- `Python/rehab_core/conflicts.py`: scheduling conflicts.
- `Python/rehab_core/replacements.py`: current single-session replacement candidates and validation.
- `Python/rehab_core/replacement_options.py`: feasible replacement times and recommendation ordering.
- `Python/rehab_core/replacement_policy.py`: provider policy + patient schedule filtering.
- `Python/rehab_core/provider_policy.py`: provider profiles, blocks, max-load and leadership overrides.
- `Python/rehab_core/daily_state.py`: effective daily session state.
- `Python/rehab_core/orphan_resolution.py`: orphan-resolution semantics.
- `Python/rehab_core/daily_statistics.py`: statistics from effective daily state.

The global multi-absence replacement optimizer is still open work. The existing single-patient ranking path must not be treated as the final global optimizer.

### Authoritative replacement facts / restore

- `Python/rehab_excel/daily_replacement_facts.py`: fact persistence model.
- `Python/rehab_excel/daily_replacement_fact_batch.py`: validate-all-before-mutate batch semantics.
- `Python/tools/restore_daily_replacement_fact.py`: single fact restore tooling.
- `Python/tools/restore_daily_replacement_facts.py`: selected/all authoritative restore backend.
- `Python/rehab_excel/daily_change_ledger_vba.py`: ledger VBA base.
- `Python/rehab_excel/daily_restore_in_session_vba.py`: in-session restore bridge.
- `Python/tools/build_daily_restore_bridge_preview.py`: native smoke/preview builder.

Authoritative truth remains `DAILY_REPLACEMENT_FACTS`; `DAILY_CHANGE_LEDGER` is only a projection/audit view.

### Daily views and MASTER

- daily render / write planning lives under `Python/rehab_excel/`.
- MASTER projection metadata uses OOXML/package reads rather than loading/saving the production workbook with openpyxl.
- `THERAPIST_DAILY`, `MASTER_SCHEDULE` and `DAILY_STATISTICS` must derive from the same effective daily truth.

### Outpatient recurring scheduling

Backend / Excel bridge ownership:

- `Python/rehab_excel/outpatient_schedule_bridge.py`
- `Python/rehab_excel/outpatient_schedule_registration.py`
- `Python/rehab_excel/outpatient_schedule_form_vba.py`
- `Python/rehab_excel/outpatient_schedule_edit_vba.py`: internal edit preload / stable target handling; public edit entry now routes through the manager.
- `Python/rehab_excel/outpatient_schedule_manager_vba.py`: business-facing list/selection UI, hidden row-to-target mapping and registration-menu routing.
- `Python/rehab_excel/outpatient_conflict_guard_vba.py`
- `Python/rehab_excel/outpatient_in_session_vba.py`
- `Python/rehab_excel/scheduling_conflict_popup_vba.py`
- `Python/tools/build_outpatient_recurring_smoke.py`: one-session combined installer including manager.

Current status:

- CREATE native green,
- stable same-row edit semantics implemented,
- application-style manager implemented in code and awaiting focused/native verification,
- old row-number InputBox flow is retired from the public edit entry,
- conflict native gate follows manager/edit verification.

`OUTPATIENT_SCHEDULE` is hidden/protected authoritative storage, not a normal user workspace.

### Registration / application shell

- registration and top-level menu/form code lives under `Python/rehab_excel/` VBA generator/installer modules.
- `frmRegistrationMenu` remains reachable from the MASTER `Διαχείριση Μητρώων` action.
- its outpatient action is now patched by `outpatient_schedule_manager_vba.py` to open the business-facing schedule manager rather than the raw create form.
- the product UX must remain application-like: MASTER/top buttons/forms, not native Excel grid navigation.
- detailed UI/product design history remains in `UI_APPLICATION_ROADMAP.md`; current cross-project UX rules are summarized in `Documentation/ONBOARDING.md` and locked in `Documentation/DECISIONS.md`.

## Closed correctness milestones

- restart-safe authoritative reconstruction: native green,
- replacement lifecycle V87-V92: native green,
- per-row / selected / restore-all authoritative restore: native green,
- daily statistics reconstruction parity: native green,
- previous outpatient combined smoke install: native green,
- outpatient CREATE: native green.

## Deferred performance work

Latest direct backend profile: `16.509s`. End-user restore latency in the low-to-mid 20-second range is accepted for now. Do not resume optimization unless there is a regression or final-polish phase.

Profiler remains at `Python/tools/profile_daily_restore_backend.py`.

## Open implementation milestones

1. focused/native verification of the application-style outpatient manager,
2. outpatient native same-entry edit gate,
3. outpatient native conflict gate,
4. broader scheduling/UI roadmap,
5. global multi-absence replacement optimizer,
6. persistent pool/robotic operating-hours configuration,
7. final restore-performance polish later.

## Business rules that must not regress

- daily facts are overlays, never destructive recurring edits,
- capacity is timeslot-based,
- existing displaced patients precede new patients,
- every orphan resolves explicitly,
- MASTER and THERAPIST_DAILY agree,
- old restore cannot overwrite newer fact history,
- bulk restore validates the whole requested set before mutation,
- baseline is immutable,
- no production/staging `.xlsm` save with openpyxl,
- outpatients count operationally but do not belong in inpatient planner UI,
- outpatients never receive robotic treatment,
- technical sheet names, worksheet rows and internal IDs are not normal end-user concepts.

## Navigation

Start: `Documentation/ONBOARDING.md`

Then use:

- exact resume: `Documentation/CURRENT_HANDOFF.md`
- durable rules: `Documentation/DECISIONS.md`
- milestone/risk state: `Documentation/PROJECT_STATUS.md`
- history: `CHANGELOG.md`

Windows: πάντα `python -m pytest`, ποτέ bare `pytest`.
