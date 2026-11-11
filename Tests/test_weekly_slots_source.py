"""Regression gate for optional lossless weekday source in an Excel workbook."""
from datetime import date, time

import pytest
from openpyxl import Workbook

from rehab_core.base_schedule import materialize_sessions_for_date
from rehab_excel.reader import read_base_schedule
from rehab_excel.weekly_slots_source import WeeklySlotsSourceError


def workbook_with_weekly_slots(tmp_path, *, weekly_rows=(), with_weekly=True):
    wb = Workbook()
    ws = wb.active
    ws.title = "PATIENT_PLANNER"
    ws.append([
        "PatientID", "Ασθενής",
        "ΦΘ_Ώρα", "ΦΘ_Ημέρες", "ΦΘ_Θεραπευτής",
        "Εργο_Ώρα", "Εργο_Ημέρες",
    ])
    ws.append(["DEMO-P1", "TEST PATIENT", time(11, 30), "Καθ/να", "T-OLD",
               time(9, 15), "Δε-Τε-Πα"])
    ws.append(["DEMO-P2", "TEST OTHER", time(10, 0), "Καθ/να", "T-OTHER",
               None, None])
    patients = wb.create_sheet("PATIENTS")
    patients.append(["PatientID", "Θάλαμος", "Ασθενής", "Μολυσματικός", "Κατάσταση"])
    patients.append(["DEMO-P1", "Α1", "TEST PATIENT", "ΟΧΙ", "ΠΑΡΩΝ"])
    patients.append(["DEMO-P2", "Α2", "TEST OTHER", "ΟΧΙ", "ΠΑΡΩΝ"])
    settings = wb.create_sheet("SETTINGS")
    settings.append(["THERAPISTS_FTH"])
    settings.append(["T-OLD"])
    if with_weekly:
        weekly = wb.create_sheet("PROGRAM_WEEKLY_SLOTS")
        weekly.append([
            "PatientID", "Treatment", "Weekday", "StartTime", "Therapist",
            "SourceCell",
        ])
        for values in weekly_rows:
            weekly.append(values)
    path = tmp_path / "weekly-source.xlsx"
    wb.save(path)
    return path


def test_no_granular_sheet_preserves_legacy_reader(tmp_path):
    path = workbook_with_weekly_slots(tmp_path, with_weekly=False)
    entries = read_base_schedule(path)
    assert len(entries) == 3
    assert any(x.patient_id == "DEMO-P1" and x.therapist_id == "T-OLD" for x in entries)


def test_granular_row_overrides_only_matching_patient_treatment(tmp_path):
    path = workbook_with_weekly_slots(
        tmp_path, weekly_rows=[
            ("DEMO-P1", "ΦΘ", "ΔΕ", time(11, 30), "T-MON", "C12"),
            ("DEMO-P1", "ΦΘ", "ΠΕ", time(12, 15), "T-THU", "C12"),
        ],
    )
    entries = read_base_schedule(path)
    assert len(entries) == 4  # P1 ΦΘ twice; P1 Εργο, P2 ΦΘ retained
    p1_fth = [x for x in entries if x.patient_id == "DEMO-P1" and x.treatment == "ΦΘ"]
    assert [(x.start_time, x.therapist_id, x.day_pattern) for x in p1_fth] == [
        (time(11, 30), "T-MON", "Δε"),
        (time(12, 15), "T-THU", "Πε"),
    ]
    monday = materialize_sessions_for_date(entries, date(2026, 11, 16))
    thursday = materialize_sessions_for_date(entries, date(2026, 11, 19))
    assert [(s.patient_id, s.therapist_id, s.start_time) for s in monday
            if s.patient_id == "DEMO-P1"] == [("DEMO-P1", "T-MON", time(11, 30))]
    assert [(s.patient_id, s.therapist_id, s.start_time) for s in thursday
            if s.patient_id == "DEMO-P1"] == [("DEMO-P1", "T-THU", time(12, 15))]
    assert any(x.treatment == "Εργο" and x.therapist_id is None for x in entries)


def test_provider_optional_other_specialty_keeps_patient_time(tmp_path):
    path = workbook_with_weekly_slots(
        tmp_path, weekly_rows=[
            ("DEMO-P1", "Εργο", "ΠΕ", time(9, 15), None, "J12"),
        ],
    )
    entries = read_base_schedule(path)
    match = [x for x in entries if x.patient_id == "DEMO-P1" and x.treatment == "Εργο"]
    assert len(match) == 1
    assert match[0].day_pattern == "Πε"
    assert match[0].therapist_id is None


@pytest.mark.parametrize("weekly_rows", [
    [("UNKNOWN", "ΦΘ", "ΔΕ", time(11, 30), "T1", "C12")],
    [("DEMO-P1", "ΦΘ", "ΔΕ", time(11, 30), None, "C12")],
    [("DEMO-P1", "ΦΘ", "ΣΑ", time(11, 30), "T1", "C12")],
    [("DEMO-P1", "ΦΘ", "ΔΕ", time(11, 30), "T1", "")],
    [
        ("DEMO-P1", "ΦΘ", "ΔΕ", time(11, 30), "T1", "C12"),
        ("DEMO-P1", "Εργο", "ΔΕ", time(11, 30), None, "J12"),
    ],
])
def test_unsafe_weekly_source_fails_closed(tmp_path, weekly_rows):
    path = workbook_with_weekly_slots(tmp_path, weekly_rows=weekly_rows)
    with pytest.raises(WeeklySlotsSourceError):
        read_base_schedule(path)


def test_audit_counts_effective_weekly_entries_not_legacy_duplicates(tmp_path):
    from rehab_excel.reader import audit_workbook
    path = workbook_with_weekly_slots(
        tmp_path, weekly_rows=[
            ("DEMO-P1", "ΦΘ", "ΔΕ", time(11, 30), "T-MON", "C12"),
            ("DEMO-P1", "ΦΘ", "ΠΕ", time(12, 15), "T-THU", "C12"),
        ],
    )
    result = audit_workbook(path)
    assert result.ok
    assert result.patient_count == 2
    assert result.base_entry_count == 4
    assert result.provider_entry_count == 3


def test_audit_blocks_unknown_weekly_identity(tmp_path):
    from rehab_excel.reader import audit_workbook
    path = workbook_with_weekly_slots(
        tmp_path, weekly_rows=[
            ("UNKNOWN", "ΦΘ", "ΔΕ", time(11, 30), "T-MON", "C12"),
        ],
    )
    result = audit_workbook(path)
    assert not result.ok
    assert any(item.code == "invalid_granular_weekly_source" for item in result.issues)
