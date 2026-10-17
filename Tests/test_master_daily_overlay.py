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


def test_replacement_overlays_master_with_effective_time_and_provider(tmp_path, monkeypatch):
    path = _book(tmp_path)
    patient = Patient(
        "P1",
        "ΜΕΝΤΗΣ",
        room="A01",
        infectious=True,
        patient_type=PatientType.INPATIENT,
    )
    entry = BaseScheduleEntry(
        base_entry_id="base-1",
        patient_id="P1",
        treatment="ΦΘ",
        start_time=time(9, 15),
        day_pattern="Δε-Τε-Πα",
        therapist_id="Σκαρώνης",
    )
    monkeypatch.setattr(
        "rehab_excel.master_daily_overlay.build_master_projection",
        lambda *_args, **_kwargs: (
            MasterProjectionRow(7, 2, "P1", "A01", "ΜΕΝΤΗΣ", True, "παρών"),
        ),
    )

    result = build_master_daily_overlay_plan(
        path,
        base_entries=[entry],
        states=[_replacement_state()],
        patients=[patient],
    )

    assert len(result.write_plan.patches) == 1
    patch = result.write_plan.patches[0]
    assert patch.sheet == "MASTER_SCHEDULE"
    assert patch.cell == "D7"
    assert patch.value == "10:45\nΔε-Τε-Πα\nΜηλιδάκης"
    assert "09:15" not in patch.value
    assert "Σκαρώνης" not in patch.value
    assert result.bindings[0].master_cell == "D7"


def test_outpatient_replacement_does_not_write_inpatient_master(tmp_path, monkeypatch):
    path = _book(tmp_path)
    patient = Patient(
        "P1",
        "ΕΞΩΤΕΡΙΚΟΣ",
        patient_type=PatientType.OUTPATIENT,
    )
    entry = BaseScheduleEntry(
        base_entry_id="base-1",
        patient_id="P1",
        treatment="ΦΘ",
        start_time=time(9, 15),
        day_pattern="Δε-Τε-Πα",
        therapist_id="Σκαρώνης",
    )
    monkeypatch.setattr(
        "rehab_excel.master_daily_overlay.build_master_projection",
        lambda *_args, **_kwargs: (),
    )

    result = build_master_daily_overlay_plan(
        path,
        base_entries=[entry],
        states=[_replacement_state()],
        patients=[patient],
    )

    assert result.write_plan.patches == ()
    assert result.bindings == ()
