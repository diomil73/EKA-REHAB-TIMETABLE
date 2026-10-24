from pathlib import Path

import pytest

from rehab_excel.outpatient_schedule_bridge import (
    OutpatientScheduleBridgeError,
    run_outpatient_schedule_bridge,
)
from rehab_excel.outpatient_schedule_registration import OutpatientSchedulePreviewReport


def test_bridge_builds_request_and_returns_json_ready_result(monkeypatch, tmp_path):
    captured = {}

    def fake_preview(
        source,
        output,
        request,
        *,
        overwrite,
        allow_therapist_double_booking,
    ):
        captured["source"] = source
        captured["output"] = output
        captured["request"] = request
        captured["overwrite"] = overwrite
        captured["allow_therapist_double_booking"] = allow_therapist_double_booking
        return OutpatientSchedulePreviewReport(
            source_path=str(source),
            output_path=str(output),
            base_entry_id="outpatient:2",
            excel_row=2,
            source_unchanged=True,
            verified_in_output=True,
            created=True,
        )

    monkeypatch.setattr(
        "rehab_excel.outpatient_schedule_bridge.create_outpatient_schedule_preview",
        fake_preview,
    )

    response = run_outpatient_schedule_bridge(
        {
            "source_path": "baseline.xlsm",
            "preview_dir": str(tmp_path),
            "overwrite": True,
            "values": {
                "patient_id": "98",
                "treatment": "ΦΘ",
                "time": "08:30",
                "days": "Δε-Τε-Πα",
                "therapist": "Αργέντος",
            },
        },
        fingerprint_service=lambda _: "def456",
    )

    request = captured["request"]
    assert request.patient_id == "98"
    assert request.treatment == "ΦΘ"
    assert request.start_time.strftime("%H:%M") == "08:30"
    assert request.day_pattern == "Δε-Τε-Πα"
    assert request.therapist_id == "Αργέντος"
    assert captured["output"] == Path(tmp_path) / "OUTPATIENT_SCHEDULE_PREVIEW.xlsm"
    assert captured["overwrite"] is True
    assert captured["allow_therapist_double_booking"] is False
    assert response["ok"] is True
    assert response["base_entry_id"] == "outpatient:2"
    assert response["source_sha256_before"] == "def456"


@pytest.mark.parametrize(
    ("payload", "message"),
    [
        ({"preview_dir": "x", "values": {}}, "source_path is required"),
        ({"source_path": "x", "values": {}}, "preview_dir is required"),
        ({"source_path": "x", "preview_dir": "y", "values": []}, "values must be an object"),
    ],
)
def test_bridge_rejects_malformed_payload(payload, message):
    with pytest.raises(OutpatientScheduleBridgeError, match=message):
        run_outpatient_schedule_bridge(payload)


def test_bridge_rejects_bad_time_before_writer():
    with pytest.raises(OutpatientScheduleBridgeError, match="time must use HH:MM"):
        run_outpatient_schedule_bridge(
            {
                "source_path": "x.xlsm",
                "preview_dir": "preview",
                "values": {
                    "patient_id": "98",
                    "treatment": "ΦΘ",
                    "time": "8.30",
                    "days": "Δε",
                },
            }
        )
