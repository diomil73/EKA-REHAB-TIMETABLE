from datetime import date, time

import pytest
from openpyxl import Workbook

from rehab_core.models import ReplacementAssignment, ReplacementProviderKind, Session
from rehab_excel.daily_replacement_facts import (
    ACTIVE_STATUS,
    RESTORED_STATUS,
    HEADERS,
    SHEET_NAME,
    DailyReplacementFactError,
    build_daily_replacement_fact_rows,
    read_daily_replacement_fact_rows,
    read_daily_replacement_facts,
)


DAY = date(2026, 10, 19)


def _session(
    session_id: str = "S1",
    *,
    patient_id: str = "P1",
    therapist_id: str = "Σάρρας",
    at: time = time(9, 15),
) -> Session:
    return Session(
        session_id=session_id,
        patient_id=patient_id,
        therapist_id=therapist_id,
        session_date=DAY,
        start_time=at,
        treatment="ΦΘ",
    )


def _replacement(
    *,
    replacement_id: str = "R1",
    target_session_id: str = "S1",
    patient_id: str = "P1",
    original_therapist_id: str = "Σάρρας",
    replacement_therapist_id: str = "Πέτσιος",
    replacement_time: time = time(10, 0),
) -> ReplacementAssignment:
    return ReplacementAssignment(
        replacement_id=replacement_id,
        target_session_id=target_session_id,
        patient_id=patient_id,
        original_therapist_id=original_therapist_id,
        replacement_therapist_id=replacement_therapist_id,
        replacement_date=DAY,
        replacement_time=replacement_time,
        reason="DAILY_INPUT therapist absence",
        replacement_provider_kind=ReplacementProviderKind.THERAPIST,
    )


def _write_fact_sheet(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    for col, header in enumerate(HEADERS, start=1):
        ws.cell(1, col).value = header
    for row_number, row in enumerate(rows, start=2):
        for col, value in enumerate(row, start=1):
            ws.cell(row_number, col).value = value
    wb.save(path)
    wb.close()


def _fact_row(
    *,
    replacement_id="R1",
    target_session_id="S1",
    patient_id="P1",
    original_therapist_id="Σάρρας",
    replacement_therapist_id="Πέτσιος",
    replacement_date=DAY,
    replacement_time="10:00",
    reason="DAILY_INPUT therapist absence",
    provider_kind="therapist",
    status=ACTIVE_STATUS,
):
    return (
        replacement_id,
        target_session_id,
        patient_id,
        original_therapist_id,
        replacement_therapist_id,
        replacement_date,
        replacement_time,
        reason,
        provider_kind,
        status,
    )


def test_build_rows_preserves_exact_replacement_semantics():
    session = _session()
    replacement = _replacement()

    rows = build_daily_replacement_fact_rows([replacement], sessions=[session])

    assert len(rows) == 1
    row = rows[0]
    assert row.replacement_id == "R1"
    assert row.target_session_id == "S1"
    assert row.patient_id == "P1"
    assert row.original_therapist_id == "Σάρρας"
    assert row.replacement_therapist_id == "Πέτσιος"
    assert row.replacement_date == DAY
    assert row.replacement_time == time(10, 0)
    assert row.provider_kind == ReplacementProviderKind.THERAPIST
    assert row.status == ACTIVE_STATUS


def test_build_rows_rejects_patient_mismatch():
    with pytest.raises(DailyReplacementFactError, match="does not match session patient"):
        build_daily_replacement_fact_rows(
            [_replacement(patient_id="OTHER")],
            sessions=[_session()],
        )


def test_build_rows_rejects_original_therapist_mismatch():
    with pytest.raises(DailyReplacementFactError, match="does not match session therapist"):
        build_daily_replacement_fact_rows(
            [_replacement(original_therapist_id="OTHER")],
            sessions=[_session()],
        )


def test_build_rows_rejects_two_active_facts_for_same_target_date():
    with pytest.raises(DailyReplacementFactError, match="Multiple active replacement facts"):
        build_daily_replacement_fact_rows(
            [
                _replacement(replacement_id="R1"),
                _replacement(
                    replacement_id="R2",
                    replacement_therapist_id="Φιλιππούσης",
                    replacement_time=time(10, 45),
                ),
            ],
            sessions=[_session()],
        )


def test_reader_missing_sheet_is_backward_compatible(tmp_path):
    path = tmp_path / "old.xlsx"
    wb = Workbook()
    wb.save(path)
    wb.close()

    assert read_daily_replacement_facts(path, sessions=[_session()], target_date=DAY) == ()


def test_reader_reconstructs_replacement_assignment(tmp_path):
    path = tmp_path / "facts.xlsx"
    _write_fact_sheet(path, [_fact_row()])

    facts = read_daily_replacement_facts(path, sessions=[_session()], target_date=DAY)

    assert facts == (_replacement(),)


def test_reader_normalizes_timezone_shifted_numeric_excel_date(tmp_path):
    path = tmp_path / "facts.xlsx"
    _write_fact_sheet(path, [_fact_row(replacement_date=46313.875)])

    facts = read_daily_replacement_facts(path, sessions=[_session()], target_date=DAY)

    assert facts == (_replacement(),)


def test_reader_defaults_blank_provider_kind_to_therapist(tmp_path):
    path = tmp_path / "facts.xlsx"
    _write_fact_sheet(path, [_fact_row(replacement_date="2026-10-19", provider_kind="")])

    facts = read_daily_replacement_facts(path, sessions=[_session()], target_date=DAY)

    assert facts[0].replacement_provider_kind == ReplacementProviderKind.THERAPIST


def test_reader_filters_other_dates_before_session_validation(tmp_path):
    path = tmp_path / "facts.xlsx"
    _write_fact_sheet(
        path,
        [
            _fact_row(
                replacement_id="OTHER-DAY",
                target_session_id="UNKNOWN",
                patient_id="P-X",
                original_therapist_id="T-X",
                replacement_therapist_id="T-Y",
                replacement_date="2026-10-20",
                reason="other day",
            ),
            _fact_row(replacement_date="2026-10-19"),
        ],
    )

    facts = read_daily_replacement_facts(path, sessions=[_session()], target_date=DAY)

    assert [fact.replacement_id for fact in facts] == ["R1"]


def test_reader_rejects_conflicting_duplicate_active_target_date(tmp_path):
    path = tmp_path / "facts.xlsx"
    _write_fact_sheet(
        path,
        [
            _fact_row(replacement_id="R1", reason="one"),
            _fact_row(
                replacement_id="R2",
                replacement_therapist_id="Φιλιππούσης",
                replacement_time="10:45",
                reason="two",
            ),
        ],
    )

    with pytest.raises(DailyReplacementFactError, match="Multiple active replacement facts"):
        read_daily_replacement_facts(path, sessions=[_session()], target_date=DAY)


def test_restored_fact_is_history_not_effective_daily_state(tmp_path):
    path = tmp_path / "facts.xlsx"
    _write_fact_sheet(path, [_fact_row(status=RESTORED_STATUS)])

    active = read_daily_replacement_facts(path, sessions=[_session()], target_date=DAY)
    history = read_daily_replacement_fact_rows(
        path,
        target_date=DAY,
        include_restored=True,
    )

    assert active == ()
    assert len(history) == 1
    assert history[0].replacement_id == "R1"
    assert history[0].status == RESTORED_STATUS


def test_restored_history_does_not_conflict_with_new_active_fact_for_same_target(tmp_path):
    path = tmp_path / "facts.xlsx"
    _write_fact_sheet(
        path,
        [
            _fact_row(replacement_id="R1", status=RESTORED_STATUS),
            _fact_row(
                replacement_id="R2",
                replacement_therapist_id="Φιλιππούσης",
                replacement_time="10:45",
                status=ACTIVE_STATUS,
            ),
        ],
    )

    active = read_daily_replacement_facts(path, sessions=[_session()], target_date=DAY)

    assert len(active) == 1
    assert active[0].replacement_id == "R2"
    assert active[0].replacement_therapist_id == "Φιλιππούσης"


def test_reader_rejects_unknown_status(tmp_path):
    path = tmp_path / "facts.xlsx"
    _write_fact_sheet(path, [_fact_row(status="MYSTERY")])

    with pytest.raises(DailyReplacementFactError, match="unsupported Status"):
        read_daily_replacement_facts(path, sessions=[_session()], target_date=DAY)


def test_reader_rejects_header_drift(tmp_path):
    path = tmp_path / "facts.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws["A1"] = "wrong"
    wb.save(path)
    wb.close()

    with pytest.raises(DailyReplacementFactError, match="headers do not match"):
        read_daily_replacement_facts(path, sessions=[_session()], target_date=DAY)
