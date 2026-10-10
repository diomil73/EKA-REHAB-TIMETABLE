from datetime import date, time
from pathlib import Path
import sys


REPO_ROOT = Path(__file__).resolve().parents[1]
TOOLS_ROOT = REPO_ROOT / "Python" / "tools"
if str(TOOLS_ROOT) not in sys.path:
    sys.path.insert(0, str(TOOLS_ROOT))

import report_therapist_day_load as report  # noqa: E402
from rehab_core.models import BaseScheduleEntry, Patient


DAY = date(2026, 10, 12)


def test_report_marks_shared_slot_and_lists_patient_treatments(monkeypatch, tmp_path):
    workbook = tmp_path / "book.xlsm"
    workbook.write_bytes(b"stub")

    entries = [
        BaseScheduleEntry("E1", "P1", "ΦΘ", time(9, 0), "Δ", "Χρήστου"),
        BaseScheduleEntry("E2", "P2", "ΦΘ", time(10, 0), "Δ", "Χρήστου"),
        BaseScheduleEntry("E3", "P3", "Ρομποτικό", time(10, 0), "Δ", "Χρήστου", robotic=True),
        BaseScheduleEntry("E4", "P4", "ΦΘ", time(11, 0), "Τρ", "Χρήστου"),
    ]
    monkeypatch.setattr(report, "read_base_schedule", lambda path: entries)
    monkeypatch.setattr(
        report,
        "read_patients",
        lambda path: [
            Patient("P1", "Ένας"),
            Patient("P2", "Δύο"),
            Patient("P3", "Τρεις"),
        ],
    )

    text = report.render_therapist_day_load_report(
        workbook,
        therapist="Χρήστου",
        target_date=DAY,
    )

    assert "Sessions: 3" in text
    assert "Occupied slots: 2" in text
    assert "10:00 (2 sessions)  <-- SHARED SLOT" in text
    assert "PatientID P2 | Δύο | ΦΘ | E2" in text
    assert "PatientID P3 | Τρεις | Ρομποτικό | E3" in text
    assert "11:00" not in text


def test_report_says_when_therapist_has_no_sessions(monkeypatch, tmp_path):
    workbook = tmp_path / "book.xlsm"
    workbook.write_bytes(b"stub")
    monkeypatch.setattr(report, "read_base_schedule", lambda path: [])
    monkeypatch.setattr(report, "read_patients", lambda path: [])

    text = report.render_therapist_day_load_report(
        workbook,
        therapist="Nobody",
        target_date=DAY,
    )

    assert "Sessions: 0" in text
    assert "Occupied slots: 0" in text
    assert "No sessions found for this therapist/date." in text
