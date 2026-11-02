import inspect

from tools import prepare_day_amendment_in_session as module


def test_in_session_preparation_reuses_authoritative_amendment_plan():
    source = inspect.getsource(module.prepare_day_amendment_in_session)

    assert "prepare_day_amendment(" in source
    assert "actor_or_source=actor_or_source" in source
    assert 'data["superseded_snapshot"]' in source
    assert 'data["replacement_snapshot"]' in source
    assert 'data["therapist_rows"]' in source
    assert 'data["started_event"]' in source
    assert 'data["finalized_event"]' in source


def test_in_session_payload_is_tsv_for_vba_and_contains_full_transition():
    source = inspect.getsource(module.prepare_day_amendment_in_session)

    assert '_record("SUPERSEDE"' in source
    assert '_record("SNAPSHOT"' in source
    assert '_record("THERAPIST"' in source
    assert '_record("EVENT"' in source
    assert '"payload_format": "day-amendment-in-session-tsv-v1"' in source
    assert '"event_rows": 2' in source


def test_in_session_payload_rejects_control_characters():
    source = inspect.getsource(module._field)

    assert '"\\t" in text' in source
    assert '"\\r" in text' in source
    assert '"\\n" in text' in source
    assert "DayAmendmentInSessionPreparationError" in source


def test_in_session_cli_is_read_prepare_only_not_excel_persistence():
    source = inspect.getsource(module)

    assert 'parser.add_argument("--input"' in source
    assert 'parser.add_argument("--payload"' in source
    assert 'parser.add_argument("--response"' in source
    assert "persist_day_amendment" not in source
    assert "Workbooks.Open" not in source
    assert "workbook.Save" not in source
