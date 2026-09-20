from datetime import time

from rehab_core.models import BaseScheduleEntry, Patient
from rehab_excel.patient_centric_preview import render_base_slot


def test_single_existing_patient_keeps_day_pattern_when_cell_becomes_temporary_multi_line():
    entry = BaseScheduleEntry(
        base_entry_id="e1",
        patient_id="p1",
        treatment="ΦΘ",
        start_time=time(9, 15),
        day_pattern="Τρ-Πε",
        therapist_id="T",
    )
    render = render_base_slot(
        "T", time(9, 15), (entry,), {"p1": Patient("p1", "ΓΚΑΛΜΑΝΗ")},
        force_day_patterns=True,
    )
    assert render.lines[0].text == "ΓΚΑΛΜΑΝΗ [Τρ-Πε]"


def test_robotic_base_patient_gets_its_own_orange_font_role():
    entry = BaseScheduleEntry(
        base_entry_id="e1",
        patient_id="p1",
        treatment="Ρομποτικό",
        start_time=time(9, 15),
        day_pattern="Τρ-Πε",
        therapist_id="T",
        robotic=True,
    )
    render = render_base_slot(
        "T", time(9, 15), (entry,), {"p1": Patient("p1", "ΜΑΚΡΙΑΔΑΚΗ")},
        force_day_patterns=True,
    )
    assert render.lines[0].text == "ΜΑΚΡΙΑΔΑΚΗ [Τρ-Πε]"
    assert render.lines[0].font_role == "robotic_orange"
