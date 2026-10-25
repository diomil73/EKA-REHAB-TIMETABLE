from rehab_excel.close_day_action_vba import CLOSE_DAY_MODULE_CODE
from rehab_excel.daily_input_action_vba import DAILY_INPUT_MODULE_CODE
from rehab_excel.daily_input_change_ledger_vba import MODULE_CODE as DAILY_INPUT_RESTORE_MODULE_CODE
from rehab_excel.global_replacement_action_vba import GLOBAL_REPLACEMENT_MODULE_CODE


PRODUCTION_ACTIONS = {
    "close_day": CLOSE_DAY_MODULE_CODE,
    "daily_input": DAILY_INPUT_MODULE_CODE,
    "daily_input_restore": DAILY_INPUT_RESTORE_MODULE_CODE,
    "global_replacements": GLOBAL_REPLACEMENT_MODULE_CODE,
}


def test_production_actions_do_not_close_or_reopen_live_workbook():
    forbidden = (
        "ThisWorkbook.Close",
        "Application.OnTime",
        "authoritative_commit_worker_cli.py",
        "post_reopen_target",
        '"reopen"',
    )
    for name, source in PRODUCTION_ACTIONS.items():
        for marker in forbidden:
            assert marker not in source, f"{name} reintroduced forbidden production handoff: {marker}"


def test_production_actions_run_python_hidden():
    for name, source in PRODUCTION_ACTIONS.items():
        assert ".Exec(" not in source, f"{name} would expose a console-backed Exec path"
        assert 'WScript.Shell").Run' in source
        assert ", 0, True)" in source


def test_production_vba_runtime_strings_do_not_contain_double_backslashes():
    for name, source in PRODUCTION_ACTIONS.items():
        assert "\\\\" not in source, (
            f"{name} contains a doubled VBA backslash. Python raw strings do not require "
            "backslash escaping for VBA string literals."
        )
