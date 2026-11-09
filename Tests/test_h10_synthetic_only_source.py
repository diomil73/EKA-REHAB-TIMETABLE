from pathlib import Path

import pytest

from tools import prepare_h10_synthetic_only_source as mod


def test_new_cb_only_and_never_overwrite(tmp_path):
    ca = tmp_path / "UNIFIED_H8_DEMO_HISTORY_CA.xlsm"
    ca.write_bytes(b"frozen old history")
    cb = tmp_path / "UNIFIED_H10_DEMO_HISTORY_CB.xlsm"
    mod.check_names(ca, cb)
    with pytest.raises(mod.DemoIsolationError, match="CB"):
        mod.check_names(ca, tmp_path / "UNIFIED_H10_DEMO_HISTORY_BZ.xlsm")
    cb.write_bytes(b"already created")
    with pytest.raises(mod.DemoIsolationError, match="CB"):
        mod.check_names(ca, cb)
    assert ca.read_bytes() == b"frozen old history"
    assert cb.read_bytes() == b"already created"


@pytest.mark.parametrize("name,column", [
    ("FINALIZED_DAY_SNAPSHOTS", 0),
    ("THERAPIST_DAY_HISTORY", 0),
    ("PATIENT_SESSION_HISTORY", 0),
    ("THERAPIST_ATTENDANCE_HISTORY", 0),
    ("DAY_HISTORY_EVENTS", 10),
])
def test_only_h8_synthetic_rows_survive(name, column):
    row = [None] * (column + 1)
    row[column] = "DEMO-H8-20260921-V1"
    assert mod.keep_row(tuple(row), id_column=column) is True
    row[column] = "finalized:2026-10-30:v2"
    assert mod.keep_row(tuple(row), id_column=column) is False
    row[column] = "DEMO-OTHER-TEST"
    with pytest.raises(mod.DemoIsolationError, match="namespace"):
        mod.keep_row(tuple(row), id_column=column)


def test_legacy_unlinked_event_is_removed_from_new_cb_not_ca():
    values = ["old-event"] + [None] * 10
    assert mod.keep_row(tuple(values), id_column=10) is False


def test_missing_synthetic_event_relation_stops():
    values = ["DEMO-INVALID"] + [None] * 10
    with pytest.raises(mod.DemoIsolationError, match="snapshot link"):
        mod.keep_row(tuple(values), id_column=10)


def test_missing_regular_snapshot_id_is_not_silently_removed():
    with pytest.raises(mod.DemoIsolationError, match="snapshot provenance"):
        mod.keep_row((None, "2026-10-30"), id_column=0)


def test_native_copy_occurs_only_after_full_read_only_classification():
    import inspect
    source = inspect.getsource(mod.install)
    assert source.index("expected = classify_rows(source)") < source.index('output.open("xb")')
    assert source.index("original_hash = sha(source)") < source.index('output.open("xb")')
    assert "verify_isolated(output, expected)" in source
    assert "ws.Rows(row_idx).Delete()" in source
    assert "ws.Protect(" in source


def test_h10_installer_cannot_reuse_old_ca():
    from tools.prepare_h10_therapist_simple_demo import H10DemoError, check_candidate
    import tempfile
    with tempfile.TemporaryDirectory() as dirname:
        folder = Path(dirname)
        ca = folder / "H8_DEMO_CA.xlsm"
        cb = folder / "H8_DEMO_CB.xlsm"
        bz = folder / "H10_DEMO_BZ.xlsm"
        ca.write_bytes(b"frozen")
        cb.write_bytes(b"isolated")
        with pytest.raises(H10DemoError, match="CB"):
            check_candidate(ca, bz)
        check_candidate(cb, bz)
