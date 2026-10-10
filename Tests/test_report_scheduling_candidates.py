from datetime import date, time
from pathlib import Path
import sys


REPO_ROOT = Path(__file__).resolve().parents[1]
TOOLS_ROOT = REPO_ROOT / "Python" / "tools"
if str(TOOLS_ROOT) not in sys.path:
    sys.path.insert(0, str(TOOLS_ROOT))

import report_scheduling_candidates as report  # noqa: E402
from rehab_core.scheduling_candidates import TherapistScheduleCandidate


def test_report_renders_ranked_candidates_and_slots(monkeypatch, tmp_path):
    workbook = tmp_path / "book.xlsm"
    workbook.write_bytes(b"stub")

    monkeypatch.setattr(
        report,
        "workbook_therapist_candidates",
        lambda *args, **kwargs: (
            TherapistScheduleCandidate(
                therapist_id="Γαύρας",
                display_name="Γαύρας",
                active_sessions=3,
                active_timeslots=3,
                max_daily_timeslots=6,
                available_slots=(time(10, 0), time(11, 0), time(12, 15)),
            ),
            TherapistScheduleCandidate(
                therapist_id="Σαρράς",
                display_name="Σαρράς",
                active_sessions=4,
                active_timeslots=4,
                max_daily_timeslots=6,
                available_slots=(time(9, 0), time(13, 0)),
            ),
        ),
    )

    text = report.render_candidate_report(
        workbook,
        patient_id="P-001",
        treatment="ΦΘ",
        target_date=date(2026, 10, 12),
    )

    assert "1. Γαύρας | load 3 | 3/6 slots" in text
    assert "   - 10:00" in text
    assert "   - 12:15" in text
    assert "2. Σαρράς | load 4 | 4/6 slots" in text
    assert "   - 13:00" in text


def test_report_says_when_no_candidate_exists(monkeypatch, tmp_path):
    workbook = tmp_path / "book.xlsm"
    workbook.write_bytes(b"stub")
    monkeypatch.setattr(report, "workbook_therapist_candidates", lambda *args, **kwargs: ())

    text = report.render_candidate_report(
        workbook,
        patient_id="P-002",
        treatment="ΦΘ",
        target_date=date(2026, 10, 12),
    )

    assert "No feasible therapist/time candidates found." in text
