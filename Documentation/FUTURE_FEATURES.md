# Future intranet and interactive timetable features

These requirements are intentionally recorded now so they remain part of the architecture even though implementation comes later.

## 1. Patient name -> electronic patient folder

The visible patient name in the timetable should act as a hyperlink/action that opens the patient's electronic folder on the hospital intranet.

Design requirements:

- Resolve the folder from stable PatientID / FolderKey, not from the visible name.
- Prefer a UNC network root such as `\\server\\share\\...` rather than a mapped drive letter.
- Keep the network root/configuration outside GitHub when it contains internal infrastructure information.
- Windows/intranet permissions remain authoritative. The workbook must not bypass network access controls.
- Moving/renaming a timetable cell must not break the patient-folder link because the link is identity-based.

## 2. Treatment time -> patient progress/evaluation chart

Selecting/clicking the treatment time should open an evaluation/progress view for the relevant patient/session.

The detailed scoring/chart rules will be specified later.

Architecture requirement now:

- progress history is keyed by PatientID (and where useful SessionID/date), never by cell address;
- changing therapist, time, day pattern, or timetable position must not detach the patient's history.

Longer-term direction:

- evolve the patient view into a chronological timeline rather than a collection of disconnected sheets/screens;
- allow the timeline to combine admission/registration, treatment plan changes, completed sessions, absences, evaluations, milestones, therapy changes and discharge/closure events;
- keep timeline events keyed to stable identities and dated operational facts so historical interpretation remains reproducible.

## 3. Therapist name -> therapist statistics dashboard

Selecting/clicking a therapist name should open a dashboard. Candidate statistics include:

- total treatments;
- daily/weekly/monthly workload;
- absences;
- replacements accepted;
- infectious-patient workload;
- robotic workload;
- available/unused timeslots;
- cancellations and patient absences affecting utilization;
- additional indicators defined later.

These statistics should be derived from assignments and operational logs rather than maintained manually.

## 4. Shared intranet / multi-user operation

The final solution will be shared on a closed hospital intranet with users having different permissions.

Design direction:

- avoid treating one shared `.xlsm` file as the only multi-user database;
- separate shared authoritative data from each user's Excel interface where practical;
- define read/write roles before production deployment;
- preserve an audit trail for operational changes;
- prevent one user's save from silently overwriting another user's changes.

Exact storage technology is intentionally undecided at this stage. The requirement is retained so current design choices do not block a later multi-user implementation.

Role-aware presentation should remain possible even before a full authentication/authorization layer exists. A future administrator, secretary/coordinator and therapist may need different default home screens, actions and levels of detail, while all roles continue to operate on the same authoritative data model.

## 5. Final app-like UI / UX

The production user interface must not look or feel like a normal Excel workbook. Excel may remain the technical host, but ordinary users should experience the system as a focused application.

Final UI goals:

- hide the Excel Ribbon and standard command surface from ordinary users;
- hide the formula bar;
- hide row numbers and column letters;
- hide worksheet gridlines;
- hide sheet tabs where navigation is provided by the application UI;
- hide unnecessary scrollbars and other spreadsheet chrome where practical;
- open users into a custom home/dashboard screen rather than a raw worksheet;
- use clear navigation, large action buttons/cards and guided forms instead of asking users to find and edit cells manually;
- keep visual language consistent across registration, daily changes, replacements, outpatient scheduling, reports and settings;
- use a modern, calm and readable design with restrained colour, clear hierarchy and accessible typography;
- reserve warning/error colours for actual operational meaning rather than decoration;
- preserve all workbook safety and authoritative-data rules behind the visual layer;
- provide a global search entry point for patients, therapists and operational records where practical;
- surface exceptions and pending work proactively instead of requiring users to discover problems by browsing sheets;
- support summary cards such as today's treatment count, current absences, unresolved replacements, incomplete documentation and other operational alerts;
- keep the scheduling/optimization machinery visually hidden: users should see decisions, explanations and actions rather than solver internals.

The preferred interaction pattern is:

`Home / Menu -> choose task -> guided form -> validation -> confirmation / preview -> apply`

rather than direct free-form spreadsheet editing.

A candidate production home screen should communicate the product as a rehabilitation operations application, not as an Excel scheduler. The working product framing is:

`Rehabilitation Operations & Scheduling System`

A useful home-screen structure to evaluate later is:

- current date and operational status;
- today's treatments / completed / pending;
- current absences and replacement status;
- unresolved operational alerts;
- Program of the Day;
- Patients;
- Therapists;
- Absences & Replacements;
- Statistics / History;
- global search.

Before the final UI is implemented, review visual and interaction ideas from comparable rehabilitation, clinic scheduling, hospital administration and therapy-practice systems. The purpose of this benchmark is primarily to learn how mature products hide operational complexity, structure dashboards, show alerts and guide users through workflows. It is not a requirement to copy their clinical or commercial feature sets.

UI/UX work is intentionally scheduled near the end of the project, after the core workflows and VBA/Python integration are stable. It will be followed by full-system testing and one integrated beta evaluation of the complete product, not fragmented beta testing of individual features.

## 6. Market-benchmark backlog additions

A review of Greek and international rehabilitation / physiotherapy / clinic systems identified several capabilities worth retaining in the architecture. The closest conceptual comparison is not a simple appointment application but a hybrid rehabilitation operations system, scheduling optimizer and patient/therapist history platform.

The following items are accepted into the future backlog because they add operational value and fit the current direction.

### 6.1 Generic room/resource model

The scheduling model should eventually be capable of treating physical rooms and equipment as first-class resources, not only special-case constraints.

Target conceptual model:

`Therapist + Patient + Treatment + Room/Resource + Time`

The existing pool, robotic and reclining constraints should be viewed as early specialized resource rules, not as the final resource abstraction. Future scheduling should be able to reject or rank assignments based on both provider availability and required room/equipment availability.

### 6.2 Therapist qualifications / service eligibility

Not every therapist should necessarily be considered interchangeable for every treatment or service.

Future provider policy should be able to express:

- qualification / competency requirements;
- service eligibility;
- equipment-specific capability;
- room/resource-specific capability where needed;
- explicit exceptions without hard-coding names into the solver.

This should extend the existing provider-policy architecture rather than create a separate scheduling path.

### 6.3 Waitlist / standby patients

A future waitlist/standby mechanism should be able to make use of capacity released by cancellations or other same-day changes.

This should integrate with the existing capacity engine rather than behave as an independent appointment list. A released slot may become a candidate opportunity for a suitable waiting patient, subject to treatment, provider, resource and patient-availability constraints.

### 6.4 Treatment-plan completion / authorized session counts

Patient history should eventually support treatment-plan progress such as:

`12 / 20 sessions completed`

Possible future fields include prescribed/authorized sessions, completed sessions, remaining sessions and expiry/end date when relevant. These counts must be derived from authoritative treatment events and not manually maintained totals.

The architecture should also remain capable of adding referral/authorization expiry warnings later, even if billing/insurance functionality is never implemented.

### 6.5 Documentation-completeness check before day closure

The future `Κλείσιμο ημέρας` workflow should not validate scheduling alone. It should be able to check whether required clinical/session documentation has been completed before the day is considered operationally complete.

A future closure summary may distinguish, for example:

- schedule resolved;
- absences resolved;
- replacements resolved;
- treatment execution recorded;
- clinical/session notes still pending.

The exact clinical documentation model will be defined separately, but the day-closure architecture must not assume that a scheduled or completed treatment automatically means its documentation is complete.

## 7. Deliberate non-goals from market comparison

Commercial clinic platforms often bundle CRM, marketing funnels, online reviews, SEO, lead management, self-service patient booking and other private-practice business features. These are not part of the core product direction unless a future operational requirement explicitly justifies them.

In particular:

- patient self-booking is not a core scheduling model for this system; complex rehabilitation scheduling remains a clinical/operational planning task;
- CRM/marketing features should not displace rehabilitation workflow work;
- market products are benchmarks for workflow clarity, information architecture and UI treatment, not a checklist that must be copied.
