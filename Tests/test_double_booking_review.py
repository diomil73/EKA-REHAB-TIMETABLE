from datetime import date, time
from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "Python"))

from rehab_core.double_booking_review import build_double_booking_review_items
from rehab_core.global_replacement_end_to_end import GlobalReplacementAssignmentResult
from rehab_core.models import Patient, ReplacementAssignment, Session, Therapist


DAY = date(2026, 9, 18)


def _session(session_id: str, patient_id: str, therapist_id: str, hh: int) -> Session:
    return Session(
        session_id=session_id,
        patient_id=patient_id,
        therapist_id=therapist_id,
        session_date=DAY,
        start_time=time(hh, 0),
        treatment="ΦΘ",
    )


def test_review_is_empty_when_normal_global_plan_has_no_unresolved_sessions():
    result = GlobalReplacementAssignmentResult(assignments=(), unresolved_request_ids=())

    assert build_double_booking_review_items(
        result,
        sessions=(),
        therapists=(),
    ) == ()


def test_review_exposes_only_real_same_slot_double_booking_options():
    target = _session("TARGET", "P1", "ABSENT", 10)
    busy_t2 = _session("BUSY-T2", "P2", "T2", 10)
    free_t3 = _session("BUSY-T3", "P3", "T3", 11)
    result = GlobalReplacementAssignmentResult(
        assignments=(), unresolved_request_ids=("TARGET",)
    )

    items = build_double_booking_review_items(
        result,
        sessions=(target, busy_t2, free_t3),
        therapists=(Therapist("T2", "Therapist Two"), Therapist("T3", "Therapist Three")),
        patients=(Patient("P1", "Patient One"),),
        timeslots=(time(10, 0),),
    )

    assert len(items) == 1
    item = items[0]
    assert item.target_session_id == "TARGET"
    assert item.patient_name == "Patient One"
    assert [(option.therapist_id, option.replacement_time) for option in item.options] == [
        ("T2", time(10, 0))
    ]
    assert "διπλοκράτηση" in item.options[0].warning


def test_review_includes_assignments_from_normal_plan_when_testing_conflicts():
    target = _session("TARGET", "P1", "ABSENT", 10)
    other_target = _session("OTHER", "P2", "ABSENT2", 10)
    result = GlobalReplacementAssignmentResult(
        assignments=(
            ReplacementAssignment(
                replacement_id="R-OTHER",
                target_session_id="OTHER",
                patient_id="P2",
                original_therapist_id="ABSENT2",
                replacement_therapist_id="T2",
                replacement_date=DAY,
                replacement_time=time(10, 0),
            ),
        ),
        unresolved_request_ids=("TARGET",),
    )

    items = build_double_booking_review_items(
        result,
        sessions=(target, other_target),
        therapists=(Therapist("T2", "T2"),),
        timeslots=(time(10, 0),),
    )

    assert len(items) == 1
    assert [(option.therapist_id, option.replacement_time) for option in items[0].options] == [
        ("T2", time(10, 0))
    ]


def test_review_keeps_unresolved_item_visible_even_when_no_exception_is_legal():
    target = _session("TARGET", "P1", "ABSENT", 10)
    result = GlobalReplacementAssignmentResult(
        assignments=(), unresolved_request_ids=("TARGET",)
    )

    items = build_double_booking_review_items(
        result,
        sessions=(target,),
        therapists=(Therapist("T2", "T2"),),
        patients=(Patient("P1", "Patient One"),),
        timeslots=(time(10, 0),),
    )

    assert len(items) == 1
    assert items[0].patient_name == "Patient One"
    assert items[0].options == ()
    assert items[0].has_options is False
