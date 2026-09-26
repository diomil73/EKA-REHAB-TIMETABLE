from rehab_core import Patient, PatientType


def test_patient_defaults_to_inpatient_for_backward_compatibility():
    patient = Patient(patient_id="P1", display_name="Ασθενής")

    assert patient.patient_type == PatientType.INPATIENT
    assert patient.is_outpatient is False
    assert patient.hospital_mrn is None


def test_outpatient_is_explicit_without_changing_identity():
    patient = Patient(
        patient_id="P100",
        display_name="Εξωτερικός ασθενής",
        patient_type=PatientType.OUTPATIENT,
        hospital_mrn="MRN-7788",
    )

    assert patient.patient_id == "P100"
    assert patient.patient_type == PatientType.OUTPATIENT
    assert patient.is_outpatient is True
    assert patient.hospital_mrn == "MRN-7788"
