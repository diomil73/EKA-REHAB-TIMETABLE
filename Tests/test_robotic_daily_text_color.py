from datetime import date, time

from rehab_core import DailySessionState, DailySessionStatus, Patient, Session
from rehab_excel import render_plan as rp


def test_robotic_daily_patient_line_uses_orange_font_role(monkeypatch):
    monkeypatch.setattr(rp, "resolve_master_schedule_rows", lambda _path, _name: (2,))
    monkeypatch.setattr(
        rp,
        "resolve_therapist_daily_cell",
        lambda _path, _provider, _slot: "B2",
    )

    target_date = date(2026, 10, 10)
    session = Session(
        "S1",
        "P1",
        "Μηλιδάκης",
        target_date,
        time(8, 30),
        treatment="Ρομποτικό",
        robotic=True,
    )
    state = DailySessionState(
        session_id="S1",
        patient_id="P1",
        session_date=target_date,
        status=DailySessionStatus.ACTIVE,
        original_therapist_id="Μηλιδάκης",
        original_time=time(8, 30),
        effective_therapist_id="Μηλιδάκης",
        effective_time=time(8, 30),
    )

    plan = rp.build_daily_excel_render_plan(
        "book.xlsm",
        [state],
        [session],
        [Patient("P1", "ΜΟΤΣΙΟ")],
    )

    assert plan.ok
    assert plan.cells[0].lines[0].font_role == rp.RenderFontRole.ROBOTIC
    assert rp.RenderFontRole.ROBOTIC.value == "robotic_orange"
