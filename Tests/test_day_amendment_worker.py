import inspect

from tools import day_amendment_worker as module


def test_worker_retries_until_workbook_is_released_and_then_reopens():
    source = inspect.getsource(module.apply_when_released)

    assert "time.monotonic()" in source
    assert "apply_day_amendment(workbook, payload)" in source
    assert "time.sleep(poll_seconds)" in source
    assert "os.startfile(str(workbook))" in source
    assert '"attempts": attempts' in source
    assert '"reopened": reopen' in source


def test_worker_fails_immediately_on_hash_mismatch():
    source = inspect.getsource(module.apply_when_released)

    assert "Workbook changed after amendment preparation" in source
    assert "raise DayAmendmentWorkerError" in source


def test_worker_cli_writes_failure_response_instead_of_mutating_request():
    source = inspect.getsource(module.main)

    assert 'parser.add_argument("--workbook"' in source
    assert 'parser.add_argument("--payload"' in source
    assert 'parser.add_argument("--response"' in source
    assert '_write_response(args.response, {"ok": False, "error": str(exc)})' in source
