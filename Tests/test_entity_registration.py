from datetime import date

from rehab_core import (
    NewPatientRequest,
    NewStudentRequest,
    NewTherapistRequest,
    Patient,
    Student,
    Therapist,
    validate_new_patient,
    validate_new_student,
    validate_new_therapist,
)


def _student(student_id: str = "S1", student_number: int = 1) -> Student:
    return Student(
        student_id=student_id,
        display_name="Μαρία Παπαδοπούλου",
        student_number=student_number,
        placement_start=date(2026, 9, 1),
        placement_end=date(2026, 12, 20),
    )


def test_new_patient_accepts_duplicate_name_but_rejects_duplicate_id():
    existing = [Patient("P1", "ΓΕΩΡΓΙΟΥ", room="101")]

    same_name = validate_new_patient(
        NewPatientRequest("P2", "ΓΕΩΡΓΙΟΥ", room="102", status="ΠΑΡΩΝ"),
        existing_patients=existing,
        allowed_rooms=["101", "102"],
        allowed_statuses=["ΠΑΡΩΝ", "ΑΔΕΙΑ"],
    )
    duplicate_id = validate_new_patient(
        NewPatientRequest(" p1 ", "ΑΛΛΟΣ"),
        existing_patients=existing,
    )

    assert same_name.allowed is True
    assert duplicate_id.allowed is False
    assert {issue.code for issue in duplicate_id.issues} == {"duplicate_patient_id"}


def test_new_patient_uses_configured_status_and_room_lists():
    check = validate_new_patient(
        NewPatientRequest("P2", "ΑΣΘΕΝΗΣ", room="999", status="ΑΓΝΩΣΤΟ"),
        allowed_rooms=["101"],
        allowed_statuses=["ΠΑΡΩΝ"],
    )

    assert check.allowed is False
    assert {issue.code for issue in check.issues} == {
        "invalid_patient_status",
        "invalid_room",
    }


def test_new_therapist_rejects_case_insensitive_duplicate_name():
    check = validate_new_therapist(
        NewTherapistRequest("  ΠΕΤΣΙΟΣ  "),
        existing_therapists=[Therapist("Πέτσιος", "Πέτσιος")],
    )

    assert check.allowed is False
    assert check.issues[0].code == "duplicate_therapist_name"


def test_new_therapist_can_validate_against_settings_name_registry():
    check = validate_new_therapist(
        NewTherapistRequest("ΝΕΟΣ ΘΕΡΑΠΕΥΤΗΣ"),
        existing_therapist_names=["Σκαρώνης", "Πέτσιος"],
    )

    assert check.allowed is True


def test_new_student_checks_identity_number_dates_and_supervisor():
    request = NewStudentRequest(
        student_id="s1",
        display_name="Νέος Φοιτητής",
        student_number=1,
        placement_start=date(2026, 11, 1),
        placement_end=date(2026, 10, 1),
        supervisor_therapist_id="UNKNOWN",
    )

    check = validate_new_student(
        request,
        existing_students=[_student()],
        known_therapist_ids=["T1"],
    )

    assert check.allowed is False
    assert {issue.code for issue in check.issues} == {
        "duplicate_student_id",
        "duplicate_student_number",
        "invalid_placement_dates",
        "unknown_supervisor",
    }


def test_new_student_allows_same_display_name_when_identity_is_unique():
    request = NewStudentRequest(
        student_id="S2",
        display_name="Μαρία Παπαδοπούλου",
        student_number=2,
        placement_start=date(2027, 1, 10),
        placement_end=date(2027, 3, 31),
        supervisor_therapist_id="T1",
    )

    check = validate_new_student(
        request,
        existing_students=[_student()],
        known_therapist_ids=["T1"],
    )

    assert check.allowed is True
