from datetime import date, time
from pathlib import Path

from openpyxl import Workbook

from rehab_core.daily_state import DailySessionState, DailySessionStatus
from rehab_core.models import BaseScheduleEntry, Patient, PatientType
from rehab_excel.master_daily_overlay import build_master_daily_overlay_plan
from rehab_excel.master_projection import MasterProjectionRow


DAY = date(2026, 10, 19)


def _book(tmp_path: Path) -> Path:
    path = tmp_path / "daily_master.xlsm"
    wb = Workbook()
    ws = wb.active
    ws.title = "MASTER_SCHEDULE"
    ws.append([
        "Μολυσματικός",
        "Θάλαμος",
        "Ασθενής",
        "ΦΘ",
        "Ρομποτικό",
        "Πισίνα",
        "Ανακλ/μενο",
        "Εργο",
        "Λογο",
        "ΕΦΑ",
        "ΨΥΧΟΛΟΓΟΙ",
        "ΑΠΟΓΕΥΜΑΤΙΝΟ ΠΡΟΓΡΑΜΜΑ",
        "Κατάσταση",
    ])
    wb.save(path)
    wb.close()
    return path


def _replacement_state() -> DailySessionState:
    return DailySessionState(
        session_id="base-1@2026-10-19",
        patient_id="P1",
        session_date=DAY,
        status=DailySessionStatus.REPLACED,
        original_therapist_id="Σκαρώνης",
        original_time=time(9, 15),
        effective_therapist_id="Μηλιδάκης",
        effective_time=time(10, 45),
        replacement_id="R1",
    )


def _cancelled_state(reason: str = "Διακομιστήριο") -> DailySessionState:
    return DailySessionState(
        session_id="base-1@2026-10-19",
        patient_id="P1",
        session_date=DAY,
        status=DailySessionStatus.CANCELLED,
        original_therapist_id="Σκαρώνης",
        original_time=time(9, 15),
        effective_therapist_id=None,
        effective_time=None,
        cancellation_id="C1",
        reason=reason,
    )


def _patient() -> Patient:
    return Patient(
        "P1",
        "ΜΕΝΤΗΣ",
        room="A01",
        infectious=True,
        patient_type=PatientType.INPATIENT,
    )


def _entry() -> BaseScheduleEntry:
    return BaseScheduleEntry(
        base_entry_id="base-1",
        patient_id="P1",
        treatment="ΦΘ",
        start_time=time(9, 15),
        day_pattern="Δε-Τε-Πα",
        therapist_id="Σκαρώνης",
    )


def _projection():
    return (
        MasterProjectionRow(7, 2, "P1", "A01", "ΜΕΝΤΗΣ", True, "παρών"),
    )


def test_replacement_overlays_master_with_effective_time_and_provider(tmp_path, monkeypatch):
    path = _book(tmp_path)
    monkeypatch.setattr(
        "rehab_excel.master_daily_overlay.build_master_projection",
        lambda *_args, **_kwargs: _projection(),
    )

    result = build_master_daily_overlay_plan(
        path,
        base_entries=[_entry()],
        states=[_replacement_state()],
        patients=[_patient()],
    )

    assert len(result.write_plan.patches) == 1
    patch = result.write_plan.patches[0]
    assert patch.sheet == "MASTER_SCHEDULE"
    assert patch.cell == "D7"
    assert patch.value == "10:45\nΔε-Τε-Πα\nΜηλιδάκης"
    assert "09:15" not in patch.value
    assert "Σκαρώνης" not in patch.value
    assert result.bindings[0].master_cell == "D7"


def test_cancelled_session_is_not_left_present_on_master(tmp_path, monkeypatch):
    path = _book(tmp_path)
    monkeypatch.setattr(
        "rehab_excel.master_daily_overlay.build_master_projection",
        lambda *_args, **_kwargs: _projection(),
    )

    result = build_master_daily_overlay_plan(
        path,
        base_entries=[_entry()],
        states=[_cancelled_state()],
        patients=[_patient()],
    )

    by_cell = {patch.cell: patch for patch in result.write_plan.patches}
    assert by_cell["D7"].value == "ΑΚΥΡΩΘΗΚΕ\nΔιακομιστήριο"
    assert by_cell["M7"].value == "ΑΚΥΡΩΘΗΚΕ"
    assert by_cell["D7"].source_tag == "master_daily_non_active_overlay"


def test_patient_absence_is_shown_as_absent_on_master(tmp_path, monkeypatch):
    path = _book(tmp_path)
    monkeypatch.setattr(
        "rehab_excel.master_daily_overlay.build_master_projection",
        lambda *_args, **_kwargs: _projection(),
    )
    state = DailySessionState(
        session_id="base-1@2026-10-19",
        patient_id="P1",
        session_date=DAY,
        status=DailySessionStatus.PATIENT_ABSENT,
        original_therapist_id="Σκαρώνης",
        original_time=time(9, 15),
        effective_therapist_id=None,
        effective_time=None,
        reason="Απουσία",
    )

    result = build_master_daily_overlay_plan(
        path,
        base_entries=[_entry()],
        states=[state],
        patients=[_patient()],
    )

    by_cell = {patch.cell: patch for patch in result.write_plan.patches}
    assert by_cell["D7"].value == "ΑΠΩΝ\nΑπουσία"
    assert by_cell["M7"].value == "ΑΠΩΝ"


def test_patient_status_remains_operational_when_another_session_is_active(tmp_path, monkeypatch):
    path = _book(tmp_path)
    monkeypatch.setattr(
        "rehab_excel.master_daily_overlay.build_master_projection",
        lambda *_args, **_kwargs: _projection(),
    )
    second_entry = BaseScheduleEntry(
        base_entry_id="base-2",
        patient_id="P1",
        treatment="Εργο",
        start_time=time(10, 0),
        day_pattern="Δε-Τε-Πα",
        therapist_id="Άλλος",
    )
    active = DailySessionState(
        session_id="base-2@2026-10-19",
        patient_id="P1",
        session_date=DAY,
        status=DailySessionStatus.ACTIVE,
        original_therapist_id="Άλλος",
        original_time=time(10, 0),
        effective_therapist_id="Άλλος",
        effective_time=time(10, 0),
    )

    result = build_master_daily_overlay_plan(
        path,
        base_entries=[_entry(), second_entry],
        states=[_cancelled_state(), active],
        patients=[_patient()],
    )

    assert "M7" not in {patch.cell for patch in result.write_plan.patches}
    assert any(
        patch.cell == "D7" and patch.value.startswith("ΑΚΥΡΩΘΗΚΕ")
        for patch in result.write_plan.patches
    )


def test_precomputed_projection_avoids_second_master_read(tmp_path, monkeypatch):
    path = _book(tmp_path)
    projection = (
        MasterProjectionRow(7, 2, "P1", "A01", "ΜΕΝΤΗΣ", False, "παρών"),
    )
    monkeypatch.setattr(
        "rehab_excel.master_daily_overlay.build_master_projection",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(AssertionError("unexpected rebuild")),
    )

    result = build_master_daily_overlay_plan(
        path,
        base_entries=[_entry()],
        states=[_replacement_state()],
        patients=[_patient()],
        master_projection_rows=projection,
    )

    assert result.write_plan.patches[0].cell == "D7"


def test_outpatient_replacement_does_not_write_inpatient_master(tmp_path, monkeypatch):
    path = _book(tmp_path)
    patient = Patient(
        "P1",
        "ΕΞΩΤΕΡΙΚΟΣ",
        patient_type=PatientType.OUTPATIENT,
    )
    monkeypatch.setattr(
        "rehab_excel.master_daily_overlay.build_master_projection",
        lambda *_args, **_kwargs: (),
    )

    result = build_master_daily_overlay_plan(
        path,
        base_entries=[_entry()],
        states=[_replacement_state()],
        patients=[patient],
    )

    assert result.write_plan.patches == ()
    assert result.bindings == ()
