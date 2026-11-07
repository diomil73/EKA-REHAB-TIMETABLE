import pytest

from tools.install_h8_single_message_refresh import (
    EXPECTED_REFRESH, FLAG, RefreshInstallError, SUCCESS_SITES, plan_patch,
)


def _sample_modules():
    result = {"Module3": "\n".join(EXPECTED_REFRESH) + "\n\nSub Independent()\nEnd Sub\n"}
    for module, lines in SUCCESS_SITES.items():
        result[module] = (
            "Public Sub Example()\n"
            + "\n".join("    " + line for line in lines)
            + "\n    MsgBox \"Επικίνδυνο σφάλμα\", vbCritical\nEnd Sub\n"
        )
    return result


def test_one_message_refresh_guards_only_success_dialogs():
    source = _sample_modules()
    patched = plan_patch(source)
    assert "Public EKA_SilentRefreshSuccess As Boolean" in patched["Module3"]
    for call in EXPECTED_REFRESH[3:9]:
        assert call in patched["Module3"]
    assert 'MsgBox "Όλα ενημερώθηκαν.", vbInformation' in patched["Module3"]
    assert "RefreshFailure:" in patched["Module3"]
    assert "Application.EnableEvents = priorEvents" in patched["Module3"]
    assert "Application.ScreenUpdating = priorScreenUpdating" in patched["Module3"]
    for module, lines in SUCCESS_SITES.items():
        for dialog in lines:
            assert f"If Not {FLAG} Then {dialog}" in patched[module]
        assert 'MsgBox "Επικίνδυνο σφάλμα", vbCritical' in patched[module]
    assert "Sub Independent()" in patched["Module3"]


def test_unexpected_refresh_contract_is_a_safety_stop():
    modules = _sample_modules()
    modules["Module3"] = modules["Module3"].replace(
        "Call RestoreReturnedTherapists", "Call UnknownRoutine"
    )
    with pytest.raises(RefreshInstallError, match="differs"):
        plan_patch(modules)


def test_duplicate_success_dialog_is_a_safety_stop():
    modules = _sample_modules()
    modules["Module1"] += (
        'Public Sub Again()\n    MsgBox "Το THERAPIST_DAILY ανανεώθηκε.", vbInformation\nEnd Sub\n'
    )
    with pytest.raises(RefreshInstallError, match="exactly one"):
        plan_patch(modules)


def test_missing_confirmation_is_a_safety_stop():
    modules = _sample_modules()
    modules["Module4"] = modules["Module4"].replace(
        'MsgBox "Κουμπιά που δημιουργήθηκαν: " & createdCount, vbInformation',
        'MsgBox "Αλλαγή σχήματος", vbInformation',
    )
    with pytest.raises(RefreshInstallError, match="exactly one"):
        plan_patch(modules)


def test_second_application_is_refused():
    patched = plan_patch(_sample_modules())
    with pytest.raises(RefreshInstallError, match="flag already exists"):
        plan_patch(patched)
