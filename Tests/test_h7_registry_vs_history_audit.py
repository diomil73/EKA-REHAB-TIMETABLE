from tools.audit_h7_registry_vs_history import _numeric_ids, REPORTED_NUMBERS


def test_reported_numbers_are_explicitly_covered():
    assert REPORTED_NUMBERS == (15, 18, 19, 20, 21, 22)


def test_numeric_registry_id_comparison_normalizes_zero_padding():
    assert _numeric_ids(["01", "1", "002", "15", "P18", ""]) == {1, 2, 15}


def test_numeric_ids_do_not_invent_sessions():
    registry = _numeric_ids(["15", "18", "19", "20", "21", "22"])
    current_final = _numeric_ids(["18", "22"])
    assert sorted(registry - current_final) == [15, 19, 20, 21]
