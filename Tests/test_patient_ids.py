from rehab_core.models import Patient
from rehab_core.patient_ids import next_sequential_patient_id


def test_next_patient_id_starts_at_one_when_no_numeric_ids_exist():
    patients = [Patient("LEGACY-A", "A"), Patient("TEST-OUT-001", "B")]
    assert next_sequential_patient_id(patients) == "1"


def test_next_patient_id_uses_highest_numeric_id_without_reuse():
    patients = [
        Patient("1", "A"),
        Patient("7", "B"),
        Patient("3", "C"),
        Patient("LEGACY-X", "D"),
    ]
    assert next_sequential_patient_id(patients) == "8"


def test_numeric_strings_with_leading_zeroes_still_advance_sequence():
    patients = [Patient("009", "A"), Patient("10", "B")]
    assert next_sequential_patient_id(patients) == "11"
