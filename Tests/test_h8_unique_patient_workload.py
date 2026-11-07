from datetime import date
from types import SimpleNamespace

import pytest

from rehab_core.daily_state import DailySessionStatus as Status
from tools.export_therapist_history_ui import (
    TherapistHistoryUiExportError,
    _unique_patient_days,
    _patient_average,
)


def session(day, session_id, patient_id, therapist, status=Status.ACTIVE):
    return SimpleNamespace(
        service_date=date.fromisoformat(day),
        session_id=session_id,
        patient_id=patient_id,
        effective_therapist_id=therapist,
        status=status,
    )


def test_unique_patients_not_sessions_and_casefolded_therapist():
    rows = (
        session("2026-10-29", "s1", "p1", "T1"),
        session("2026-10-29", "s2", "p1", "T1"),
        session("2026-10-29", "s3", "p2", "t1", Status.REPLACED),
        session("2026-10-29", "s4", "p3", None, Status.CANCELLED),
    )
    assert _unique_patient_days(rows) == {(date(2026, 10, 29), "t1"): 2}


def test_average_is_over_recorded_finalized_therapist_days():
    counts = _unique_patient_days((
        session("2026-10-29", "s1", "p1", "T1"),
        session("2026-10-29", "s2", "p2", "T1"),
        session("2026-10-30", "s3", "p3", "T1"),
    ))
    report = SimpleNamespace(
        therapist_id="T1",
        period_start=date(2026, 10, 1),
        period_end=date(2026, 10, 31),
        reported_days=2,
    )
    assert _patient_average(report, counts, "monthly") == "1.50"


def test_duplicate_current_final_session_or_unassigned_delivery_stops():
    duplicated = session("2026-10-30", "s1", "p1", "T1")
    with pytest.raises(TherapistHistoryUiExportError, match="Duplicate"):
        _unique_patient_days((duplicated, duplicated))
    with pytest.raises(TherapistHistoryUiExportError, match="effective therapist"):
        _unique_patient_days((session("2026-10-30", "s1", "p1", None),))


def test_therapist_vba_load_uses_exact_patient_count_and_hides_unavailable_absences():
    from rehab_excel.therapist_history_report_vba import THERAPIST_HISTORY_MODULE_CODE as src

    assert 'If UBound(parts) <> 14' in src
    assert 'ws.Cells(rowIndex, 14).Interior.Color' in src
    assert 'WorkloadTrafficColor(CDbl(ws.Cells(rowIndex, 14).Value2))' in src
    assert "loadPerDay = uniquePatientsPerDay" in src
    assert 'ws.ChartObjects("ekaTherapistAttendanceChart").Delete' in src
    assert "Μη διαθέσιμο ιστορικό απουσιών / εργασίμων" in src
