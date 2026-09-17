"""Core scheduling logic for the Rehab Center system."""

from .availability import is_patient_available, is_therapist_available
from .conflicts import (
    ConflictKind,
    OperationalOccurrence,
    ScheduleConflict,
    build_operational_occurrences,
    find_operational_conflicts,
)
from .day_patterns import (
    RehabWeekday,
    parse_day_pattern,
    patterns_are_complementary,
    patterns_overlap,
)
from .models import (
    AbsenceKind,
    DailyAbsence,
    Patient,
    ReplacementAssignment,
    Session,
    Therapist,
)
from .replacements import (
    ReplacementCandidate,
    create_replacement_assignment,
    find_replacement_candidates,
)

__all__ = [
    "AbsenceKind",
    "ConflictKind",
    "DailyAbsence",
    "OperationalOccurrence",
    "Patient",
    "RehabWeekday",
    "ReplacementAssignment",
    "ReplacementCandidate",
    "ScheduleConflict",
    "Session",
    "Therapist",
    "build_operational_occurrences",
    "create_replacement_assignment",
    "find_operational_conflicts",
    "find_replacement_candidates",
    "is_patient_available",
    "is_therapist_available",
    "parse_day_pattern",
    "patterns_are_complementary",
    "patterns_overlap",
]
