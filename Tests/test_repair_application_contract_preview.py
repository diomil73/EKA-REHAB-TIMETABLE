import inspect

from tools import repair_application_contract_preview as repair_tool


def test_contract_repair_is_narrow_and_uses_existing_manager_installer():
    assert repair_tool.ALLOWED_MISSING == {
        repair_tool.OUTPATIENT_MANAGER_MODULE,
        repair_tool.OUTPATIENT_MANAGER_FORM,
    }
    source = inspect.getsource(repair_tool.repair_preview)
    assert "install_outpatient_schedule_manager" in source
    assert "unexpected required component(s) are missing" in source
    assert "assert_application_contract(path)" in source


def test_contract_repair_has_rollback_copy():
    source = inspect.getsource(repair_tool.repair_preview)
    assert "shutil.copy2(path, backup_path)" in source
    assert "shutil.copy2(backup_path, path)" in source
