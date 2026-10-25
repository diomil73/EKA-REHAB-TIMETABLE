from rehab_excel.close_day_action_vba import CLOSE_DAY_MODULE_CODE
from rehab_excel.daily_input_action_vba import DAILY_INPUT_MODULE_CODE


PRODUCTION_ACTIONS = {
    "close_day": CLOSE_DAY_MODULE_CODE,
    "daily_input": DAILY_INPUT_MODULE_CODE,
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
