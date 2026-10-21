from datetime import date, time

from rehab_core.global_replacement_end_to_end import build_global_replacement_assignments
from rehab_core.models import (
    AbsenceKind,
    DailyAbsence,
    ReplacementProviderKind,
    Session,
    Therapist,
)
from rehab_core.service_hours import ROBOTIC_SERVICE_ID, ServiceOperatingHours, SpecialServiceHoursBook


TODAY = date(2026, 10, 21)


def _session(session_id: str, patient_id: str, therapist_id: str, hh: int, *, robotic: bool = False) -> Session:
    return Session(
        session_id=session_id,
        patient_id=patient_id,
        therapist_id=therapist_id,
        session_date=TODAY,
        start_time=time(hh, 0),
        robotic=robotic,
    )


def test_builds_authoritative_replacement_assignments_from_global_plan():
    targets = (
        _session("A", "P-A", "ABSENT-1", 9),
        _session("B", "P-B", "ABSENT-2", 9),
    )
    absences = (
        DailyAbsence(AbsenceKind.THERAPIST, "ABSENT-1", TODAY),
        DailyAbsence(AbsenceKind.THERAPIST, "ABSENT-2", TODAY),
    )

    result = build_global_replacement_assignments(
        targets,
        therapists=(Therapist("T1", "One"), Therapist("T2", "Two")),
        sessions=targets,
        absences=absences,
        timeslots=(time(9, 0), time(10, 0)),
        replacement_id_factory=lambda session_id: f"R-{session_id}",
    )

    assert result.unresolved_request_ids == ()
    assert len(result.assignments) == 2
    by_target = {item.target_session_id: item for item in result.assignments}
    assert by_target["A"].replacement_id == "R-A"
    assert by_target["B"].replacement_id == "R-B"
    assert {item.replacement_therapist_id for item in result.assignments} == {"T1", "T2"}
    assert all(item.replacement_provider_kind == ReplacementProviderKind.THERAPIST for item in result.assignments)


def test_service_hours_filter_before_global_assignment():
    target = _session("ROB", "P-R", "ABSENT", 9, robotic=True)
    hours = SpecialServiceHoursBook(
        services={
            ROBOTIC_SERVICE_ID: ServiceOperatingHours(
                ROBOTIC_SERVICE_ID,
                {TODAY.weekday(): (time(10, 0),)},
            )
        }
    )

    result = build_global_replacement_assignments(
        (target,),
        therapists=(Therapist("T1", "Robot", robotic_capable=True),),
        sessions=(target,),
        absences=(DailyAbsence(AbsenceKind.THERAPIST, "ABSENT", TODAY),),
        timeslots=(time(9, 0), time(10, 0)),
        service_hours=hours,
        replacement_id_factory=lambda session_id: f"R-{session_id}",
    )

    assert result.unresolved_request_ids == ()
    assert result.assignments[0].replacement_time == time(10, 0)


def test_closed_special_service_leaves_request_unresolved_without_fake_assignment():
    target = _session("ROB", "P-R", "ABSENT", 9, robotic=True)
    hours = SpecialServiceHoursBook(
        services={
            ROBOTIC_SERVICE_ID: ServiceOperatingHours(
                ROBOTIC_SERVICE_ID,
                {TODAY.weekday(): ()},
            )
        }
    )

    result = build_global_replacement_assignments(
        (target,),
        therapists=(Therapist("T1", "Robot", robotic_capable=True),),
        sessions=(target,),
        absences=(DailyAbsence(AbsenceKind.THERAPIST, "ABSENT", TODAY),),
        timeslots=(time(9, 0), time(10, 0)),
        service_hours=hours,
        replacement_id_factory=lambda session_id: f"R-{session_id}",
    )

    assert result.assignments == ()
    assert result.unresolved_request_ids == ("ROB",)
