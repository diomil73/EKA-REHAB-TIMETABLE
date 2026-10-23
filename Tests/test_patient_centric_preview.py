from datetime import time

from rehab_core.models import BaseScheduleEntry, Patient
from rehab_excel.patient_centric_preview import render_base_slot


def entry(entry_id, patient_id, days, treatment="ΦΘ", robotic=False):
    return BaseScheduleEntry(
        base_entry_id=entry_id,
        patient_id=patient_id,
        treatment=treatment,
        start_time=time(8, 30),
        day_pattern=days,
        therapist_id="T1",
        robotic=robotic,
    )


def test_pair_shows_each_patient_with_own_days():
    patients = {
        "A": Patient("A", "ΓΚΑΛΜΑΝΗ"),
        "B": Patient("B", "ΠΕΤΙΡΟΠΟΥΛΟΣ"),
    }
    render = render_base_slot(
        "T1",
        time(8, 30),
        [entry("a", "A", "Τρ-Πε"), entry("b", "B", "Δε-Τε-Πα")],
        patients,
    )
    assert render.text == "ΓΚΑΛΜΑΝΗ [Τρ-Πε]\nΠΕΤΙΡΟΠΟΥΛΟΣ [Δε-Τε-Πα]"


def test_same_patient_split_shows_name_once_and_keeps_assignment_details():
    patients = {"M": Patient("M", "ΜΟΤΣΙΟΣ")}
    render = render_base_slot(
        "T1",
        time(8, 30),
        [
            entry("m1", "M", "Δε-Τε-Πα", "Ρομποτικό", True),
            entry("m2", "M", "Τρ-Πε", "ΦΘ"),
        ],
        patients,
    )
    assert render.text == "ΜΟΤΣΙΟΣ\nΔε-Τε-Πα ΡΟΜΠ | Τρ-Πε ΦΘ"
    assert render.text.count("ΜΟΤΣΙΟΣ") == 1
