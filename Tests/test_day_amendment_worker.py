import inspect

from tools import day_amendment_worker as module


def test_worker_waits_for_exclusive_file_release_before_excel_apply_and_reopen():
    source = inspect.getsource(module.apply_when_released)

    assert "time.monotonic()" in source
    assert "_workbook_is_exclusively_available(workbook)" in source
    assert "time.sleep(poll_seconds)" in source
    assert "apply_day_amendment(workbook, payload)" in source
    assert source.index("_workbook_is_exclusively_available(workbook)") < source.index(
        "apply_day_amendment(workbook, payload)"
    )
    assert "os.startfile(str(workbook))" in source
    assert '"attempts": probes' in source
    assert '"reopened": reopen' in source


def test_worker_file_lock_probe_uses_exclusive_windows_handle_without_excel():
    source = inspect.getsource(module._workbook_is_exclusively_available)

    assert "win32file.CreateFile" in source
    assert "win32con.GENERIC_READ | win32con.GENERIC_WRITE" in source
    assert "OPEN_EXISTING" in source
    assert "Workbooks.Open" not in source
    assert "apply_day_amendment" not in source


def test_worker_file_lock_probe_treats_sharing_and_lock_violations_as_still_locked():
    source = inspect.getsource(module._workbook_is_exclusively_available)

    assert "except Exception as exc" in source
    assert "_windows_error_code(exc) in (32, 33)" in source
    assert "return False" in source

    helper = inspect.getsource(module._windows_error_code)
    assert 'getattr(exc, "winerror", None)' in helper
    assert "exc.args[0]" in helper


def test_worker_can_refresh_payload_after_release_before_strict_apply():
    source = inspect.getsource(module.apply_when_released)

    assert "refresh_payload_after_release" in source
    assert "prepare_day_amendment(" in source
    assert "actor_or_source=actor_or_source" in source
    assert source.index("prepare_day_amendment(") < source.index(
        "apply_day_amendment(workbook, payload)"
    )
    assert '"payload_refreshed_after_release": refresh_payload_after_release' in source


def test_worker_fails_immediately_on_hash_mismatch_after_release():
    source = inspect.getsource(module.apply_when_released)

    assert "Workbook changed after amendment preparation" in source
    assert "raise DayAmendmentWorkerError" in source


def test_worker_cli_supports_post_release_refresh_and_writes_failure_response():
    source = inspect.getsource(module.main)

    assert 'parser.add_argument("--workbook"' in source
    assert 'parser.add_argument("--payload"' in source
    assert 'parser.add_argument("--response"' in source
    assert 'parser.add_argument("--refresh-payload-after-release"' in source
    assert 'parser.add_argument("--actor-or-source"' in source
    assert "refresh_payload_after_release=args.refresh_payload_after_release" in source
    assert "actor_or_source=args.actor_or_source" in source
    assert '_write_response(args.response, {"ok": False, "error": str(exc)})' in source
