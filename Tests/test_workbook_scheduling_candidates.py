from datetime import date, time

import rehab_excel.workbook_scheduling_candidates as workbook_candidates
from rehab_core.models import BaseScheduleEntry, Patient, Student
from rehab_excel.reader import WorkbookSettings


DAY = date(2026, 10, 12)  # Monday


def _settings() -> WorkbookSettings:
    return WorkbookSettings(
        therapist_names=("Σαρράς", "Γαύρας"),
        standard_timeslots=(time(9, 0), time(10, 0), time(11, 0)),
        reclined_timeslots=(time(9, 30),),
        day_patterns=("Καθ/να",),
        therapies=("ΦΘ", "Ρομποτικό"),
        patient_statuses=("Ενεργός",),
        yes_no_values=("Ν", "Ο"),
        rooms=("A1",),
    )


def _no_students(path):
    return []


def test_materialize_recurring_sessions_only_for_target_weekday():
    entries = [
        BaseScheduleEntry("MON", "P1", "ΦΘ", time(9, 0), "Δ", "Σαρράς"),
        BaseScheduleEntry("TUE", "P2", "ΦΘ", time(10, 0), "Τρ", "Γαύρας"),
        BaseScheduleEntry("NOPROVIDER", "P3", "Ανακλινόμενο", time(11, 0), "Δ", None),
    ]

    sessions = workbook_candidates.materialize_recurring_sessions(entries, DAY)

    assert [session.session_id for session in sessions] == ["MON"]
    assert sessions[0].therapist_id == "Σαρράς"


def test_workbook_candidates_use_real_settings_and_recurring_program(monkeypatch):
    entries = [
        BaseScheduleEntry("S1", "P1", "ΦΘ", time(9, 0), "Καθ/να", "Σαρράς"),
        BaseScheduleEntry("S2", "P2", "ΦΘ", time(10, 0), "Καθ/να", "Σαρράς"),
        BaseScheduleEntry("S3", "P3", "ΦΘ", time(9, 0), "Καθ/να", "Γαύρας"),
    ]
    monkeypatch.setattr(workbook_candidates, "read_settings", lambda path: _settings())
    monkeypatch.setattr(workbook_candidates, "read_base_schedule", lambda path: entries)
    monkeypatch.setattr(workbook_candidates, "read_students", _no_students)
    monkeypatch.setattr(
        workbook_candidates,
        "read_patients",
        lambda path: [Patient("P-NEW", "Νέος"), Patient("P1", "Ένας"), Patient("P2", "Δύο"), Patient("P3", "Τρεις")],
    )

    candidates = workbook_candidates.workbook_therapist_candidates(
        "dummy.xlsm",
        patient_id="P-NEW",
        treatment="ΦΘ",
        target_date=DAY,
    )

    assert [candidate.display_name for candidate in candidates] == ["Γαύρας", "Σαρράς"]
    assert candidates[0].available_slots == (time(10, 0), time(11, 0))
    assert candidates[1].available_slots == (time(11, 0),)


def test_existing_robotic_assignment_marks_provider_robotic_capable(monkeypatch):
    entries = [
        BaseScheduleEntry(
            "R1",
            "P1",
            "Ρομποτικό",
            time(9, 0),
            "Τρ",
            "Γαύρας",
            robotic=True,
        )
    ]
    monkeypatch.setattr(workbook_candidates, "read_settings", lambda path: _settings())
    monkeypatch.setattr(workbook_candidates, "read_base_schedule", lambda path: entries)
    monkeypatch.setattr(workbook_candidates, "read_students", _no_students)
    monkeypatch.setattr(workbook_candidates, "read_patients", lambda path: [Patient("P-NEW", "Νέος")])

    candidates = workbook_candidates.workbook_therapist_candidates(
        "dummy.xlsm",
        patient_id="P-NEW",
        treatment="Ρομποτικό",
        target_date=DAY,
    )

    assert [candidate.display_name for candidate in candidates] == ["Γαύρας"]


def test_additional_confirmed_robotic_provider_can_be_supplied(monkeypatch):
    monkeypatch.setattr(workbook_candidates, "read_settings", lambda path: _settings())
    monkeypatch.setattr(workbook_candidates, "read_base_schedule", lambda path: [])
    monkeypatch.setattr(workbook_candidates, "read_students", _no_students)
    monkeypatch.setattr(workbook_candidates, "read_patients", lambda path: [Patient("P-NEW", "Νέος")])

    candidates = workbook_candidates.workbook_therapist_candidates(
        "dummy.xlsm",
        patient_id="P-NEW",
        treatment="Ρομποτικό",
        target_date=DAY,
        additional_robotic_capable_names=("Σαρράς",),
    )

    assert [candidate.display_name for candidate in candidates] == ["Σαρράς"]


def test_students_are_excluded_from_physio_candidate_pool(monkeypatch):
    settings = WorkbookSettings(
        therapist_names=("Σαρράς", "φοιτ 1", "Γαύρας"),
        standard_timeslots=(time(9, 0),),
        reclined_timeslots=(),
        day_patterns=("Καθ/να",),
        therapies=("ΦΘ",),
        patient_statuses=(),
        yes_no_values=(),
        rooms=(),
    )
    student = Student(
        "ST1",
        "φοιτ 1",
        date(2026, 10, 1),
        date(2026, 12, 31),
    )
    monkeypatch.setattr(workbook_candidates, "read_settings", lambda path: settings)
    monkeypatch.setattr(workbook_candidates, "read_base_schedule", lambda path: [])
    monkeypatch.setattr(workbook_candidates, "read_students", lambda path: [student])
    monkeypatch.setattr(workbook_candidates, "read_patients", lambda path: [Patient("P-NEW", "Νέος")])

    candidates = workbook_candidates.workbook_therapist_candidates(
        "dummy.xlsm",
        patient_id="P-NEW",
        treatment="ΦΘ",
        target_date=DAY,
    )

    assert [candidate.display_name for candidate in candidates] == ["Γαύρας", "Σαρράς"]


def test_unknown_patient_id_is_rejected(monkeypatch):
    monkeypatch.setattr(workbook_candidates, "read_settings", lambda path: _settings())
    monkeypatch.setattr(workbook_candidates, "read_base_schedule", lambda path: [])
    monkeypatch.setattr(workbook_candidates, "read_students", _no_students)
    monkeypatch.setattr(workbook_candidates, "read_patients", lambda path: [Patient("P1", "Ένας")])

    try:
        workbook_candidates.workbook_therapist_candidates(
            "dummy.xlsm",
            patient_id="DOES-NOT-EXIST",
            treatment="ΦΘ",
            target_date=DAY,
        )
    except ValueError as exc:
        assert "Unknown PatientID" in str(exc)
    else:
        raise AssertionError("Expected ValueError for unknown PatientID")


def test_non_physio_treatment_is_not_silently_routed_through_physio_engine(monkeypatch):
    monkeypatch.setattr(workbook_candidates, "read_settings", lambda path: _settings())

    try:
        workbook_candidates.workbook_therapist_candidates(
            "dummy.xlsm",
            patient_id="P-NEW",
            treatment="Εργο",
            target_date=DAY,
        )
    except ValueError as exc:
        assert "physiotherapist scheduling engine" in str(exc)
    else:
        raise AssertionError("Expected ValueError for non-physiotherapy treatment")
