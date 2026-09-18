from datetime import date, time
from pathlib import Path

import pytest

from rehab_core import (
    DailySessionState,
    DailySessionStatus,
    Patient,
    ReplacementProviderKind,
    Session,
    Student,
)
from rehab_excel import render_plan as rp

DAY = date(2026, 9, 18)


@pytest.fixture
def fake_layout(monkeypatch):
    def master_rows(_path, name):
        return {
            "ΑΣΘΕΝΗΣ Α": (2,),
            "ΑΣΘΕΝΗΣ Β": (3,),
            "ΑΣΘΕΝΗΣ Γ": (4,),
        }.get(name, ())

    mapping = {
        ("Θ1", time(12, 15)): "B7",
        ("Θ2", time(12, 15)): "C7",
        ("Θ2", time(13, 0)): "C8",
        ("φοιτ 1", time(13, 0)): "H18",
        ("ΜΑΡΙΑ ΦΟΙΤΗΤΡΙΑ", time(13, 0)): "H18",
    }

    def daily_cell(_path, provider, slot):
        try:
            return mapping[(provider, slot)]
        except KeyError as exc:
            raise rp.LayoutSafetyError(f"No THERAPIST_DAILY cell for {provider}") from exc

    monkeypatch.setattr(rp, "resolve_master_schedule_rows", master_rows)
    monkeypatch.setattr(rp, "resolve_therapist_daily_cell", daily_cell)


def patient(pid="P1", name="ΑΣΘΕΝΗΣ Α", infectious=False):
    return Patient(pid, name, infectious=infectious)


def session(sid="S1", pid="P1", therapist="Θ1", at=time(12, 15), treatment="ΦΘ", robotic=False):
    return Session(sid, pid, therapist, DAY, at, treatment=treatment, robotic=robotic)


def state(status=DailySessionStatus.ACTIVE, **kwargs):
    values = dict(
        session_id="S1",
        patient_id="P1",
        session_date=DAY,
        status=status,
        original_therapist_id="Θ1",
        original_time=time(12, 15),
        effective_therapist_id="Θ1" if status == DailySessionStatus.ACTIVE else None,
        effective_time=time(12, 15) if status == DailySessionStatus.ACTIVE else None,
    )
    values.update(kwargs)
    return DailySessionState(**values)


def test_active_state_maps_real_master_and_daily_cells(fake_layout):
    plan = rp.build_daily_excel_render_plan(
        "book.xlsm", [state()], [session()], [patient()]
    )
    assert plan.ok
    assert plan.bindings[0].master_schedule_cell == "D2"
    assert plan.bindings[0].original_daily_cell == "B7"
    assert plan.bindings[0].effective_daily_cell is None
    assert plan.cells[0].text == "ΑΣΘΕΝΗΣ Α"
    assert plan.cells[0].lines[0].strike_through is False


def test_patient_absence_strikes_original_line(fake_layout):
    plan = rp.build_daily_excel_render_plan(
        "book.xlsm",
        [state(DailySessionStatus.PATIENT_ABSENT)],
        [session()],
        [patient()],
    )
    line = plan.cells[0].lines[0]
    assert line.strike_through is True
    assert line.font_role == rp.RenderFontRole.MUTED


def test_therapist_absence_keeps_original_cell_but_strikes(fake_layout):
    plan = rp.build_daily_excel_render_plan(
        "book.xlsm",
        [state(DailySessionStatus.THERAPIST_ABSENT)],
        [session()],
        [patient()],
    )
    assert plan.bindings[0].effective_daily_cell is None
    assert plan.cells[0].lines[0].strike_through is True


def test_replacement_same_time_maps_to_new_provider_cell(fake_layout):
    replaced = state(
        DailySessionStatus.REPLACED,
        effective_therapist_id="Θ2",
        effective_time=time(12, 15),
        effective_provider_kind=ReplacementProviderKind.THERAPIST,
        replacement_id="R1",
    )
    plan = rp.build_daily_excel_render_plan(
        "book.xlsm", [replaced], [session()], [patient()]
    )
    binding = plan.bindings[0]
    assert binding.original_daily_cell == "B7"
    assert binding.effective_daily_cell == "C7"
    cells = {item.cell: item for item in plan.cells}
    assert cells["B7"].lines[0].strike_through is True
    assert cells["C7"].lines[0].role == rp.RenderLineRole.REPLACEMENT
    assert cells["C7"].lines[0].strike_through is False


def test_replacement_new_time_maps_to_new_timeslot(fake_layout):
    replaced = state(
        DailySessionStatus.REPLACED,
        effective_therapist_id="Θ2",
        effective_time=time(13, 0),
        effective_provider_kind=ReplacementProviderKind.THERAPIST,
        replacement_id="R1",
    )
    plan = rp.build_daily_excel_render_plan(
        "book.xlsm", [replaced], [session()], [patient()]
    )
    assert plan.bindings[0].effective_daily_cell == "C8"


def test_student_replacement_can_resolve_legacy_student_column(fake_layout):
    student = Student(
        student_id="ST1",
        display_name="ΜΑΡΙΑ ΦΟΙΤΗΤΡΙΑ",
        student_number=1,
        placement_start=date(2026, 9, 1),
        placement_end=date(2026, 12, 31),
    )
    replaced = state(
        DailySessionStatus.REPLACED,
        effective_therapist_id="ST1",
        effective_time=time(13, 0),
        effective_provider_kind=ReplacementProviderKind.STUDENT,
        replacement_id="R1",
    )
    plan = rp.build_daily_excel_render_plan(
        "book.xlsm", [replaced], [session()], [patient()], students=[student]
    )
    assert plan.ok
    assert plan.bindings[0].effective_daily_cell == "H18"


def test_infectious_patient_uses_yellow_fill_and_border(fake_layout):
    plan = rp.build_daily_excel_render_plan(
        "book.xlsm", [state()], [session()], [patient(infectious=True)]
    )
    cell = plan.cells[0]
    assert cell.fill_role == rp.RenderFillRole.INFECTIOUS
    assert cell.border_role == rp.RenderBorderRole.INFECTIOUS


def test_robotic_infectious_keeps_pink_fill_and_yellow_border(fake_layout):
    plan = rp.build_daily_excel_render_plan(
        "book.xlsm",
        [state()],
        [session(robotic=True, treatment="Ρομποτικό")],
        [patient(infectious=True)],
    )
    cell = plan.cells[0]
    assert cell.fill_role == rp.RenderFillRole.ROBOTIC
    assert cell.border_role == rp.RenderBorderRole.INFECTIOUS
    assert plan.bindings[0].master_schedule_cell == "E2"


def test_multiple_patients_in_same_daily_cell_are_aggregated(fake_layout):
    s1 = session()
    s2 = session(sid="S2", pid="P2")
    st1 = state()
    st2 = DailySessionState(
        session_id="S2",
        patient_id="P2",
        session_date=DAY,
        status=DailySessionStatus.ACTIVE,
        original_therapist_id="Θ1",
        original_time=time(12, 15),
        effective_therapist_id="Θ1",
        effective_time=time(12, 15),
    )
    plan = rp.build_daily_excel_render_plan(
        "book.xlsm",
        [st1, st2],
        [s1, s2],
        [patient(), patient("P2", "ΑΣΘΕΝΗΣ Β")],
    )
    assert len(plan.cells) == 1
    assert plan.cells[0].text == "ΑΣΘΕΝΗΣ Α\nΑΣΘΕΝΗΣ Β"


def test_missing_patient_is_reported_not_guessed(fake_layout):
    plan = rp.build_daily_excel_render_plan("book.xlsm", [state()], [session()], [])
    assert not plan.ok
    assert plan.issues[0].code == "MISSING_PATIENT"


def test_ambiguous_master_row_is_reported(monkeypatch, fake_layout):
    monkeypatch.setattr(rp, "resolve_master_schedule_rows", lambda _p, _n: (2, 9))
    plan = rp.build_daily_excel_render_plan(
        "book.xlsm", [state()], [session()], [patient()]
    )
    assert not plan.ok
    assert plan.issues[0].code == "MASTER_PATIENT_ROW_AMBIGUOUS"


def test_unsupported_treatment_is_reported(fake_layout):
    plan = rp.build_daily_excel_render_plan(
        "book.xlsm", [state()], [session(treatment="ΑΛΛΟ")], [patient()]
    )
    assert not plan.ok
    assert plan.issues[0].code == "UNSUPPORTED_TREATMENT"


def test_states_from_more_than_one_date_are_rejected(fake_layout):
    other = DailySessionState(
        session_id="S2",
        patient_id="P2",
        session_date=date(2026, 9, 19),
        status=DailySessionStatus.ACTIVE,
        original_therapist_id="Θ1",
        original_time=time(12, 15),
        effective_therapist_id="Θ1",
        effective_time=time(12, 15),
    )
    with pytest.raises(ValueError, match="one date"):
        rp.build_daily_excel_render_plan(
            "book.xlsm",
            [state(), other],
            [session(), session(sid="S2", pid="P2")],
            [patient(), patient("P2", "ΑΣΘΕΝΗΣ Β")],
        )
