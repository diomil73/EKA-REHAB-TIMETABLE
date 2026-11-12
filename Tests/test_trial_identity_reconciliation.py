"""Privacy-safe patient identity matching contract. Synthetic names only."""
from rehab_excel.trial_identity_reconciliation import (
    PatientIdentity, OldPatientIdentity, normalize_patient_name,
    suggest_identity_matches,
)


def test_unique_exact_unicode_name_produces_only_unconfirmed_candidate():
    source = [PatientIdentity(12, "Α", "  ΔΟΚΙΜΑΣΤΙΚΟΣ   ΑΣΘΕΝΗΣ ")]
    old = [OldPatientIdentity("ID-21", "δοκιμαστικοσ ασθενησ")]
    result = suggest_identity_matches(source, old)
    assert len(result) == 1
    assert result[0].suggested_patient_id == "ID-21"
    assert result[0].state == "EXACT_NAME_CANDIDATE_UNCONFIRMED"


def test_new_person_does_not_inherit_unrelated_old_id():
    candidates = suggest_identity_matches(
        [PatientIdentity(4, "Α", "ΑΛΛΟΣ ΑΣΘΕΝΗΣ")],
        [OldPatientIdentity("ID-OLD", "ΠΑΛΙΟΣ ΑΣΘΕΝΗΣ")],
    )
    assert candidates[0].state == "NO_EXACT_OLD_MATCH"
    assert candidates[0].suggested_patient_id is None


def test_duplicate_current_names_cannot_share_one_old_id():
    result = suggest_identity_matches(
        [PatientIdentity(4, "Α", "ΙΔΙΟ ΟΝΟΜΑ"),
         PatientIdentity(22, "Β", "ΙΔΙΟ ΟΝΟΜΑ")],
        [OldPatientIdentity("ID-21", "ΙΔΙΟ ΟΝΟΜΑ")],
    )
    assert all(x.state == "AMBIGUOUS_CURRENT_NAME" for x in result)
    assert all(x.suggested_patient_id is None for x in result)


def test_duplicate_old_names_require_manual_identity_resolution():
    result = suggest_identity_matches(
        [PatientIdentity(4, "Α", "ΙΔΙΟ ΟΝΟΜΑ")],
        [OldPatientIdentity("ID-11", "ΙΔΙΟ ΟΝΟΜΑ"),
         OldPatientIdentity("ID-12", "ΙΔΙΟ ΟΝΟΜΑ")],
    )
    assert result[0].state == "AMBIGUOUS_OLD_IDENTITY"
    assert result[0].suggested_patient_id is None


def test_reused_old_id_blocked_even_with_unique_names():
    result = suggest_identity_matches(
        [PatientIdentity(4, "Α", "ΟΝΟΜΑ ΕΝΑ")],
        [OldPatientIdentity("ID-11", "ΟΝΟΜΑ ΕΝΑ"),
         OldPatientIdentity("ID-11", "ΟΝΟΜΑ ΔΥΟ")],
    )
    assert result[0].state == "AMBIGUOUS_OLD_IDENTITY"
    assert result[0].suggested_patient_id is None


def test_duplicate_source_location_is_never_trusted():
    result = suggest_identity_matches(
        [PatientIdentity(4, "Α", "ΟΝΟΜΑ ΕΝΑ"),
         PatientIdentity(4, "Α", "ΟΝΟΜΑ ΔΥΟ")],
        [OldPatientIdentity("ID-11", "ΟΝΟΜΑ ΕΝΑ")],
    )
    assert all(x.state == "DUPLICATE_SOURCE_LOCATION" for x in result)


def test_normalization_handles_final_sigma_accents_and_space():
    assert normalize_patient_name("  ΆΣΘΕΝΗΣ   ΔΟΚΙΜΗ  ") == normalize_patient_name(
        "ασθενησ δοκιμη"
    )



def test_ward_or_doctor_reassignment_does_not_change_exact_id_candidate():
    # Source location identifies a snapshot row, never the patient's identity.
    source = [PatientIdentity(38, "Β", "ΣΥΝΘΕΤΙΚΟΣ ΑΣΘΕΝΗΣ")]
    history = [OldPatientIdentity("PERMANENT-ID-12", "ΣΥΝΘΕΤΙΚΟΣ ΑΣΘΕΝΗΣ")]
    suggestion = suggest_identity_matches(source, history)[0]
    assert suggestion.suggested_patient_id == "PERMANENT-ID-12"
    assert suggestion.state == "EXACT_NAME_CANDIDATE_UNCONFIRMED"
