from datetime import date, time

from rehab_core.models import AbsenceKind, DailyAbsence, Patient, Session, Therapist
from rehab_core.therapist_absence_queue import build_therapist_absence_replacement_queue

DAY = date(2026, 9, 18)
SLOTS = (time(8, 30), time(9, 15), time(10, 0))


def session(session_id, patient_id, therapist_id, at):
    return Session(
        session_id=session_id,
        patient_id=patient_id,
        therapist_id=therapist_id,
        session_date=DAY,
        start_time=at,
        treatment="ΦΘ",
    )


def test_therapist_absence_creates_one_queue_item_per_affected_patient():
    sessions = [
        session("S1", "P1", "T-ABS", time(8, 30)),
        session("S2", "P2", "T-ABS", time(9, 15)),
        session("S3", "P3", "T-OTHER", time(10, 0)),
    ]
    absences = [
        DailyAbsence(AbsenceKind.THERAPIST, "T-ABS", DAY),
    ]
    queue = build_therapist_absence_replacement_queue(
        sessions=sessions,
        absences=absences,
        therapists=[Therapist("T-ABS", "Absent"), Therapist("T-FREE", "Free")],
        patients=[Patient("P1", "One"), Patient("P2", "Two"), Patient("P3", "Three")],
        timeslots=SLOTS,
    )
    assert [item.session_id for item in queue.items] == ["S1", "S2"]
    assert all(item.options for item in queue.items)
    assert all(item.options[0].provider_id == "T-FREE" for item in queue.items)


def test_patient_absence_same_session_is_not_sent_to_replacement_queue():
    sessions = [session("S1", "P1", "T-ABS", time(8, 30))]
    absences = [
        DailyAbsence(AbsenceKind.THERAPIST, "T-ABS", DAY),
        DailyAbsence(AbsenceKind.PATIENT, "P1", DAY),
    ]
    queue = build_therapist_absence_replacement_queue(
        sessions=sessions,
        absences=absences,
        therapists=[Therapist("T-ABS", "Absent"), Therapist("T-FREE", "Free")],
        patients=[Patient("P1", "One")],
        timeslots=SLOTS,
    )
    assert queue.items == ()
    assert queue.skipped_patient_absent == 1


def test_another_absent_therapist_is_not_suggested():
    sessions = [session("S1", "P1", "T-A", time(8, 30))]
    absences = [
        DailyAbsence(AbsenceKind.THERAPIST, "T-A", DAY),
        DailyAbsence(AbsenceKind.THERAPIST, "T-B", DAY),
    ]
    queue = build_therapist_absence_replacement_queue(
        sessions=sessions,
        absences=absences,
        therapists=[Therapist("T-A", "A"), Therapist("T-B", "B"), Therapist("T-C", "C")],
        patients=[Patient("P1", "One")],
        timeslots=SLOTS,
    )
    assert [option.provider_id for option in queue.items[0].options] == ["T-C"]
