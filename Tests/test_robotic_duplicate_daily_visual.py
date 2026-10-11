from datetime import date, time

from rehab_core import DailySessionState, DailySessionStatus, Patient, Session
from rehab_excel import render_plan as rp
from rehab_excel.render_writeback import build_daily_write_plan


DAY = date(2026, 10, 12)


def _state(session_id: str) -> DailySessionState:
    return DailySessionState(
        session_id=session_id,
        patient_id="P1",
        session_date=DAY,
        status=DailySessionStatus.ACTIVE,
        original_therapist_id="ΜΗΛΙΔΑΚΗΣ",
        original_time=time(12, 15),
        effective_therapist_id="ΜΗΛΙΔΑΚΗΣ",
        effective_time=time(12, 15),
    )


def test_same_patient_twice_keeps_only_robotic_line_orange(monkeypatch):
    monkeypatch.setattr(rp, "resolve_master_schedule_rows", lambda _path, _name: (2,))
    monkeypatch.setattr(
        rp,
        "resolve_therapist_daily_cell",
        lambda _path, _provider, _slot: "B7",
    )

    patient = Patient("P1", "ΜΟΤΣΙΟΣ")
    sessions = [
        Session(
            "S1",
            "P1",
            "ΜΗΛΙΔΑΚΗΣ",
            DAY,
            time(12, 15),
            treatment="ΦΘ",
            robotic=False,
        ),
        Session(
            "S2",
            "P1",
            "ΜΗΛΙΔΑΚΗΣ",
            DAY,
            time(12, 15),
            treatment="Ρομποτικό",
            robotic=False,
        ),
    ]

    plan = rp.build_daily_excel_render_plan(
        "book.xlsm",
        [_state("S1"), _state("S2")],
        sessions,
        [patient],
    )

    assert plan.ok
    assert len(plan.cells) == 1
    cell = plan.cells[0]
    assert cell.cell == "B7"
    assert cell.text == "ΜΟΤΣΙΟΣ\nΜΟΤΣΙΟΣ"
    assert cell.lines[0].font_role == rp.RenderFontRole.DEFAULT
    assert cell.lines[1].font_role == rp.RenderFontRole.ROBOTIC
    assert cell.fill_role == rp.RenderFillRole.ROBOTIC

    write_plan = build_daily_write_plan(plan)
    patch = write_plan.patches[0]
    assert patch.text_runs[0].font_role == "default"
    assert patch.text_runs[1].font_role == "robotic_orange"
    assert patch.text_runs[1].start == len("ΜΟΤΣΙΟΣ") + 2
