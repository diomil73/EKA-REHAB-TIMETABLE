from rehab_excel.therapist_registration import choose_therapist_target_row


def test_therapist_target_row_follows_last_real_registry_value():
    assert choose_therapist_target_row(24) == 25


def test_therapist_target_row_never_overwrites_settings_header():
    assert choose_therapist_target_row(1) == 2


def test_therapist_target_row_handles_empty_registry_upper_bound():
    assert choose_therapist_target_row(0) == 2
