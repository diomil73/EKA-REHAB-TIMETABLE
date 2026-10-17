from __future__ import annotations

from datetime import date, time

import pytest

from rehab_core.models import (
    DailySessionCancellation,
    Patient,
    Session,
    SessionCancellationKind,
)
from rehab_core.orphan_resolution import create_treatment_suspension
from rehab_excel.daily_operational_fact_writer import (
    DailyOperationalFactWriteError,
    TREATMENT_SUSPENSION_STATUS,
    build_treatment_suspension_rows,
)


DAY = date(2026, 10, 19)


def _session() -> Session:
    return Session(
        session_id="S1",
        patient_id="P1",
        therapist_id="T1",
        session_date=DAY,
        start_time=time(9, 15),
        treatment="ΦΘ",
    )


def test_build_treatment_suspension_row_preserves_patient_and_original_time():
    session = _session()
    suspension = create_treatment_suspension(session)

    rows = build_treatment_suspension_rows(
        [suspension],
        sessions=[session],
        patients=[Patient("P1", "ΜΕΝΤΗΣ")],
    )

    assert len(rows) == 1
    row = rows[0]
    assert row.target_session_id == "S1"
    assert row.patient_name == "ΜΕΝΤΗΣ"
    assert row.start_time == time(9, 15)
    assert row.status == TREATMENT_SUSPENSION_STATUS


def test_build_treatment_suspension_rows_rejects_other_cancellation_kind():
    session = _session()
    cancellation = DailySessionCancellation(
        cancellation_id="C1",
        target_session_id="S1",
        cancellation_date=DAY,
        kind=SessionCancellationKind.PATIENT_NO_SHOW,
        reason="ΔΕΝ ΠΡΟΣΗΛΘΕ",
    )

    with pytest.raises(DailyOperationalFactWriteError, match="not treatment suspension"):
        build_treatment_suspension_rows(
            [cancellation],
            sessions=[session],
            patients=[Patient("P1", "ΜΕΝΤΗΣ")],
        )


def test_build_treatment_suspension_rows_rejects_date_mismatch():
    session = _session()
    suspension = DailySessionCancellation(
        cancellation_id="SUSP",
        target_session_id="S1",
        cancellation_date=date(2026, 10, 20),
        kind=SessionCancellationKind.TREATMENT_SUSPENDED,
        reason="Αναστολή θεραπείας",
    )

    with pytest.raises(DailyOperationalFactWriteError, match="does not match session date"):
        build_treatment_suspension_rows(
            [suspension],
            sessions=[session],
            patients=[Patient("P1", "ΜΕΝΤΗΣ")],
        )
