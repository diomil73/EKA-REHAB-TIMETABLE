# Workbook Sheets

Τα 13 φύλλα που επιβεβαιώνονται από τον immutable baseline validator είναι:

- PATIENT_PLANNER
- MASTER_SCHEDULE
- THERAPIST_ATTENDANCE
- REPLACEMENTS
- REPLACEMENT_LOG
- DAILY_LOG
- STATISTICS
- PATIENTS
- NEW_PATIENT
- SETTINGS
- SESSIONS
- CONFLICT_LOG
- THERAPIST_DAILY

## Schema evolution

Το νέο `STUDENTS` ορίζεται ως authoritative registry για τους φοιτητές. Δεν θεωρείται μέρος του immutable baseline των 13 φύλλων. Προστίθεται μόνο μέσω explicit preview/schema migration σε αντίγραφο του `.xlsm` και, αφού επιβεβαιωθεί, θα χρησιμοποιείται από το formal student-registration workflow.

Authoritative columns A:H:

- StudentID
- Φοιτητής
- StudentNumber
- PlacementStart
- PlacementEnd
- SupervisorTherapist
- ReplacementCapable
- RoboticCapable

Workbooks χωρίς `STUDENTS` παραμένουν αναγνώσιμα και επιστρέφουν κενό student registry.

## Working classification

### Source / input oriented
- PATIENTS
- SETTINGS
- PATIENT_PLANNER
- STUDENTS (after explicit schema migration)
- NEW_PATIENT

### Daily operations
- THERAPIST_ATTENDANCE
- REPLACEMENTS
- REPLACEMENT_LOG
- DAILY_LOG
- CONFLICT_LOG

### Views / reports
- MASTER_SCHEDULE
- THERAPIST_DAILY
- STATISTICS

### Structured session data
- SESSIONS

Η ταξινόμηση αυτή επιβεβαιώνεται sheet-by-sheet πριν αλλάξει συμπεριφορά.
