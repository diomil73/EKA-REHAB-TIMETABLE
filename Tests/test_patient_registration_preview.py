from rehab_excel.patient_registration import resolve_infectious_cell_value


def test_infectious_true_uses_configured_yes_value():
    assert resolve_infectious_cell_value(True, ("ΝΑΙ", "ΟΧΙ")) == "ΝΑΙ"


def test_infectious_false_uses_configured_no_value():
    assert resolve_infectious_cell_value(False, ("ΝΑΙ", "ΟΧΙ")) == "ΟΧΙ"


def test_infectious_true_has_safe_reader_compatible_fallback():
    assert resolve_infectious_cell_value(True, ()) == "Ν"


def test_infectious_false_can_fall_back_to_blank():
    assert resolve_infectious_cell_value(False, ()) == ""
