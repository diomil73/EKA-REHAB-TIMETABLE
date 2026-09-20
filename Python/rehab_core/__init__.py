"""Core scheduling logic for the Rehab Center system."""

from .availability import (
    is_patient_available,
    is_student_available,
    is_therapist_available,
)
from .base_schedule import materialize_sessions_for_date, pattern_applies_on_date
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
    BaseScheduleEntry,
    DailyAbsence,
    Patient,
    ReplacementAssignment,
    ReplacementProviderKind,
    Session,
    Student,
    StudentAssignment,
    Therapist,
)
from .replacements import (
    ReplacementCandidate,
    create_replacement_assignment,
    find_replacement_candidates,
)
from .students import (
    StudentDisplayState,
    build_student_display_map,
    student_display_state,
    students_for_session,
)
from .workload import (
    StudentWorkload,
    TherapistWorkload,
    calculate_student_workload,
    calculate_therapist_workload,
)

__all__ = [
    "AbsenceKind",
    "BaseScheduleEntry",
    "ConflictKind",
    "DailyAbsence",
    "DailySessionStatus",
    "DailySessionState",
    "OperationalOccurrence",
    "Patient",
    "RehabWeekday",
    "ReplacementAssignment",
    "ReplacementCandidate",
    "ReplacementProviderKind",
    "ScheduleConflict",
    "Session",
    "Student",
    "StudentAssignment",
    "StudentDisplayState",
    "StudentWorkload",
    "Therapist",
    "TherapistWorkload",
    "build_daily_session_states",
    "build_operational_occurrences",
    "build_student_display_map",
    "calculate_student_workload",
    "calculate_therapist_workload",
    "create_replacement_assignment",
    "find_operational_conflicts",
    "find_replacement_candidates",
    "is_patient_available",
    "is_student_available",
    "is_therapist_available",
    "materialize_sessions_for_date",
    "parse_day_pattern",
    "patterns_are_complementary",
    "patterns_overlap",
    "pattern_applies_on_date",
    "student_display_state",
    "students_for_session",
]
from .therapist_absence_queue import (
    TherapistAbsenceReplacementItem,
    TherapistAbsenceReplacementQueue,
    build_therapist_absence_replacement_queue,
)

