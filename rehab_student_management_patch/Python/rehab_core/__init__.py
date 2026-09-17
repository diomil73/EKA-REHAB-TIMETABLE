"""Core scheduling logic for the Rehab Center system."""

from .availability import is_patient_available, is_therapist_available
from .conflicts import (
    ConflictKind,
    OperationalOccurrence,
    ScheduleConflict,
    build_operational_occurrences,
    find_operational_conflicts,
)
from .daily_state import (
    DailySessionState,
    DailySessionStatus,
    build_daily_session_states,
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
    Student,
    StudentAssignment,
    Therapist,
)
from .workload import TherapistWorkload, calculate_therapist_workload
from .students import (
    StudentDisplayState,
    build_student_display_map,
    student_display_state,
    students_for_session,
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
    "DailySessionStatus",
    "DailySessionState",
    "OperationalOccurrence",
    "Patient",
    "RehabWeekday",
    "ReplacementAssignment",
    "ReplacementCandidate",
    "ScheduleConflict",
    "Session",
    "Student",
    "StudentAssignment",
    "StudentDisplayState",
    "Therapist",
    "TherapistWorkload",
    "build_daily_session_states",
    "build_operational_occurrences",
    "build_student_display_map",
    "calculate_therapist_workload",
    "create_replacement_assignment",
    "find_operational_conflicts",
    "find_replacement_candidates",
    "is_patient_available",
    "is_therapist_available",
    "parse_day_pattern",
    "patterns_are_complementary",
    "patterns_overlap",
    "student_display_state",
    "students_for_session",
]
