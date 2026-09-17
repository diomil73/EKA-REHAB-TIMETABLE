# Changelog

## Unreleased

### Added
- Replacement ranking based first on operational daily workload, then exact-time match, then another common free timeslot.
- Alternative-timeslot visibility for replacement candidates.
- Student replacement providers with active-placement and capability checks.
- Student daily capacity of 5 patient/timeslots.
- Student workload and exact-timeslot availability engine.
- Provider-kind metadata so therapist and student replacements remain distinct.
- Student double-booking conflict classification.

### Changed
- A provider busy at the requested time is no longer automatically removed from replacement suggestions when another common free timeslot exists.
- Infectious workload is now a tie-breaker after total load and exact-time availability, rather than overriding the confirmed load-first priority.

### Unchanged
- Base schedule remains immutable.
- `Rehab_Center_System_v27_1.xlsm` baseline remains untouched.
