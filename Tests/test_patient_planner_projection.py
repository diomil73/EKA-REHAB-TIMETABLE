from rehab_excel.patient_planner_projection import (
    is_outpatient_type,
    planner_identity_formulas,
)


def test_planner_identity_formulas_keep_blank_source_cells_blank():
    formulas = planner_identity_formulas(12)
    assert formulas == (
        '=IF(PATIENTS!C12="","",PATIENTS!A12)',
        '=IF(PATIENTS!C12="","",PATIENTS!B12)',
        '=IF(PATIENTS!C12="","",PATIENTS!C12)',
        '=IF(PATIENTS!C12="","",PATIENTS!D12)',
        '=IF(PATIENTS!C12="","",PATIENTS!E12)',
    )


def test_planner_identity_formulas_reject_header_row():
    try:
        planner_identity_formulas(1)
    except ValueError as exc:
        assert ">= 2" in str(exc)
    else:
        raise AssertionError("Expected row 1 to be rejected")


def test_outpatient_type_marker_recognizes_greek_and_english():
    assert is_outpatient_type("Εξωτερικός") is True
    assert is_outpatient_type("OUTPATIENT") is True
    assert is_outpatient_type("Εσωτερικός") is False
