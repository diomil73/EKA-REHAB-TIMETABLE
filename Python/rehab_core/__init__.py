"""Core scheduling logic for the Rehab Center system."""

from .capacity import (
    DailyCapacityStatus,
    RecurringCapacityCheck,
    RecurringCapacityIssue,
    student_daily_capacity,
    therapist_daily_capacity,
    validate_new_patient_assignment_capacity,
    validate_recurring_timeslot_capacity,
)
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
from .student_transition import (
    StudentEndAffectedEntry,
    StudentEndReassignmentOption,
    StudentEndTransitionPlan,
    expired_student_label,
    plan_student_end_transition,
    resolve_student_base_entries,
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
    "DailyCapacityStatus",
    "RecurringCapacityCheck",
    "RecurringCapacityIssue",
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
    "StudentEndAffectedEntry",
    "StudentEndReassignmentOption",
    "StudentEndTransitionPlan",
    "StudentWorkload",
    "Therapist",
    "TherapistWorkload",
    "build_daily_session_states",
    "build_operational_occurrences",
    "build_student_display_map",
    "calculate_student_workload",
    "calculate_therapist_workload",
    "create_replacement_assignment",
    "expired_student_label",
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
    "plan_student_end_transition",
    "resolve_student_base_entries",
    "student_display_state",
    "student_daily_capacity",
    "students_for_session",
    "therapist_daily_capacity",
    "validate_new_patient_assignment_capacity",
    "validate_recurring_timeslot_capacity",
]
