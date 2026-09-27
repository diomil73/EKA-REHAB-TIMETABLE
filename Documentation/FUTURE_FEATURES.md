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
- preserve all workbook safety and authoritative-data rules behind the visual layer.

The preferred interaction pattern is:

`Home / Menu -> choose task -> guided form -> validation -> confirmation / preview -> apply`

rather than direct free-form spreadsheet editing.

Before the final UI is implemented, review visual and interaction ideas from comparable systems such as clinic scheduling dashboards, hospital administration software, booking systems, internal ERP tools and polished Excel-based applications. These references are for inspiration only; the final design should be tailored to this rehabilitation workflow rather than copied.

UI/UX work is intentionally scheduled near the end of the project, after the core workflows and VBA/Python integration are stable. It will be followed by full-system testing and one integrated beta evaluation of the complete product, not fragmented beta testing of individual features.
