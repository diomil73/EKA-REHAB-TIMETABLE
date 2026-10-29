from datetime import time

from rehab_core.models import BaseScheduleEntry, Patient, PatientType
from rehab_excel.recurring_daily_style import _roles_for_lines


def test_same_patient_split_schedule_styles_only_robotic_line_orange():
    patient = Patient("1", "ΜΟΤΣΙΟΣ")
    entries = [
        BaseScheduleEntry(
            base_entry_id="planner:2:ΦΘ",
            patient_id="1",
            treatment="ΦΘ",
            start_time=time(8, 30),
            day_pattern="Τρ-Πε",
            therapist_id="Μηλιδάκης",
            robotic=False,
        ),
        BaseScheduleEntry(
            base_entry_id="planner:2:Ρομποτικό",
            patient_id="1",
            treatment="Ρομποτικό",
            start_time=time(8, 30),
            day_pattern="Δε-Τε-Πα",
            therapist_id="Μηλιδάκης",
            robotic=True,
        ),
    ]

    text = "ΜΟΤΣΙΟΣ [Τρ-Πε]\nΜΟΤΣΙΟΣ [Δε-Τε-Πα]"
    runs = _roles_for_lines(text, entries, {"1": patient})

    assert len(runs) == 2
    assert runs[0].font_role == "default"
    assert runs[1].font_role == "robotic_orange"
    assert runs[0].start == 1
    assert runs[1].start == len("ΜΟΤΣΙΟΣ [Τρ-Πε]") + 2


def test_mixed_recurring_cell_keeps_outpatient_line_neutral_black():
    motsios = Patient("1", "ΜΟΤΣΙΟΣ")
    outpatient = Patient(
        "99",
        "τεστ",
        patient_type=PatientType.OUTPATIENT,
    )
    entries = [
        BaseScheduleEntry(
            base_entry_id="planner:2:ΦΘ",
            patient_id="1",
            treatment="ΦΘ",
            start_time=time(8, 30),
            day_pattern="Τρ-Πε",
            therapist_id="Μηλιδάκης",
            robotic=False,
        ),
        BaseScheduleEntry(
            base_entry_id="planner:2:Ρομποτικό",
            patient_id="1",
            treatment="Ρομποτικό",
            start_time=time(8, 30),
            day_pattern="Δε-Τε-Πα",
            therapist_id="Μηλιδάκης",
            robotic=True,
        ),
        BaseScheduleEntry(
            base_entry_id="planner:99:ΦΘ",
            patient_id="99",
            treatment="ΦΘ",
            start_time=time(8, 30),
            day_pattern="Καθ/να",
            therapist_id="Μηλιδάκης",
            robotic=False,
        ),
    ]

    text = "ΜΟΤΣΙΟΣ [Τρ-Πε]\nΜΟΤΣΙΟΣ [Δε-Τε-Πα]\nτεστ"
    runs = _roles_for_lines(
        text,
        entries,
        {"1": motsios, "99": outpatient},
    )

    assert [run.font_role for run in runs] == [
        "default",
        "robotic_orange",
        "default",
    ]
