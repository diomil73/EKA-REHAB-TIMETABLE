from datetime import date, time
from pathlib import Path

from openpyxl import Workbook

from rehab_core.daily_state import DailySessionState, DailySessionStatus
from rehab_core.models import Patient, ReplacementProviderKind, Session
from rehab_excel import operational_overlay as overlay_mod
from rehab_excel.operational_overlay import build_operational_overlay_write_plan


def _book(tmp_path: Path) -> Path:
    path = tmp_path / "overlay.xlsx"
    wb = Workbook()
    wb.active.title = "THERAPIST_DAILY"
    wb.create_sheet("MASTER_SCHEDULE")
    wb.save(path)
    return path


def _session(session_id="s1", patient_id="p1", therapist="T1", slot=time(12, 15), robotic=False):
    return Session(
        session_id=session_id,
        patient_id=patient_id,
        therapist_id=therapist,
        session_date=date(2026, 9, 18),
        start_time=slot,
        treatment="ΦΘ",
        robotic=robotic,
    )


def _replacement_state(session_id="s1", patient_id="p1", original="T1", replacement="T2"):
    return DailySessionState(
        session_id=session_id,
        patient_id=patient_id,
        session_date=date(2026, 9, 18),
        status=DailySessionStatus.REPLACED,
        original_therapist_id=original,
        original_time=time(12, 15),
        effective_therapist_id=replacement,
        effective_time=time(12, 15),
        replacement_id="r1",
        effective_provider_kind=ReplacementProviderKind.THERAPIST,
    )


def test_replacement_is_italic_not_struck_and_preserves_target_text(tmp_path, monkeypatch):
    path = _book(tmp_path)
    monkeypatch.setattr(
        overlay_mod,
        "_read_daily_cells",
        lambda _: (
            {"B7": "ΖΑΛΟΚΩΣΤΑΣ", "G17": "ΑΛΛΟΣ ΑΣΘΕΝΗΣ"},
            {"B7": ("infectious_yellow", "infectious_yellow"), "G17": (None, None)},
        ),
    )
    monkeypatch.setattr(
        overlay_mod,
        "_resolve_provider_cell",
        lambda _path, provider_id, _slot, **_: "B7" if provider_id == "T1" else "G17",
    )

    result = build_operational_overlay_write_plan(
        path,
        [_replacement_state()],
        [_session()],
        [Patient("p1", "ΖΑΛΟΚΩΣΤΑΣ", infectious=True)],
        provider_labels={"T2": "Φιλιππούσης"},
    )

    patches = {patch.cell: patch for patch in result.write_plan.patches}
    original = patches["B7"]
    target = patches["G17"]

    assert original.value == "ΖΑΛΟΚΩΣΤΑΣ\n→ Φιλιππούσης 12:15"
    assert original.text_runs[0].strike_through is False
    assert original.text_runs[0].italic is True
    assert target.value == "ΑΛΛΟΣ ΑΣΘΕΝΗΣ\nΖΑΛΟΚΩΣΤΑΣ"
    assert target.fill_role == "infectious_yellow"
    assert target.border_role == "infectious_yellow"


def test_patient_absence_strikes_original_line(tmp_path, monkeypatch):
    path = _book(tmp_path)
    monkeypatch.setattr(
        overlay_mod,
        "_read_daily_cells",
        lambda _: ({"B7": "ΖΑΛΟΚΩΣΤΑΣ"}, {"B7": (None, None)}),
    )
    monkeypatch.setattr(overlay_mod, "_resolve_provider_cell", lambda *args, **kwargs: "B7")
    state = DailySessionState(
        session_id="s1",
        patient_id="p1",
        session_date=date(2026, 9, 18),
        status=DailySessionStatus.PATIENT_ABSENT,
        original_therapist_id="T1",
        original_time=time(12, 15),
        effective_therapist_id=None,
        effective_time=None,
    )
    result = build_operational_overlay_write_plan(
        path,
        [state],
        [_session()],
        [Patient("p1", "ΖΑΛΟΚΩΣΤΑΣ")],
    )
    patch = result.write_plan.patches[0]
    assert patch.value == "ΖΑΛΟΚΩΣΤΑΣ"
    assert patch.text_runs[0].strike_through is True
    assert patch.text_runs[0].italic is False


def test_two_replacements_can_share_one_target_cell(tmp_path, monkeypatch):
    path = _book(tmp_path)
    monkeypatch.setattr(
        overlay_mod,
        "_read_daily_cells",
        lambda _: (
            {"B7": "ΑΣΘΕΝΗΣ Α", "C7": "ΑΣΘΕΝΗΣ Β", "G17": "ΥΠΑΡΧΩΝ"},
            {"B7": (None, None), "C7": (None, None), "G17": (None, None)},
        ),
    )

    def resolver(_path, provider_id, _slot, **_):
        return {"T1": "B7", "T3": "C7", "T2": "G17"}[provider_id]

    monkeypatch.setattr(overlay_mod, "_resolve_provider_cell", resolver)
    states = [
        _replacement_state("s1", "p1", "T1", "T2"),
        _replacement_state("s2", "p2", "T3", "T2"),
    ]
    sessions = [
        _session("s1", "p1", "T1"),
        _session("s2", "p2", "T3"),
    ]
    patients = [Patient("p1", "ΑΣΘΕΝΗΣ Α"), Patient("p2", "ΑΣΘΕΝΗΣ Β")]
    result = build_operational_overlay_write_plan(path, states, sessions, patients)
    target = {patch.cell: patch for patch in result.write_plan.patches}["G17"]
    assert target.value == "ΥΠΑΡΧΩΝ\nΑΣΘΕΝΗΣ Α\nΑΣΘΕΝΗΣ Β"
    assert len(result.bindings) == 2


def test_replaced_state_property_is_not_struck():
    state = _replacement_state()
    assert state.original_should_be_struck_through is False
