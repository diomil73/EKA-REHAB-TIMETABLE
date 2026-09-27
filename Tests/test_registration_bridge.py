from pathlib import Path

import pytest

from rehab_excel.registration_bridge import RegistrationBridgeError, run_registration_bridge
from rehab_excel.registration_orchestrator import RegistrationKind, RegistrationPreviewResult


def test_bridge_builds_outpatient_request_and_returns_json_ready_result(tmp_path):
    captured = {}

    def fake_preview(source, preview_dir, request, *, overwrite):
        captured["source"] = source
        captured["preview_dir"] = preview_dir
        captured["request"] = request
        captured["overwrite"] = overwrite
        return RegistrationPreviewResult(
            kind=RegistrationKind.PATIENT,
            subject_key="98",
            display_name=request.display_name,
            output_path=str(Path(preview_dir) / "NEW_PATIENT_PREVIEW.xlsm"),
            excel_row=99,
            source_unchanged=True,
            verified_in_output=True,
        )

    response = run_registration_bridge(
        {
            "source_path": "baseline.xlsm",
            "preview_dir": str(tmp_path),
            "action": "new_patient",
            "overwrite": True,
            "values": {
                "patient_id": "",
                "patient_type": "Εξωτερικός",
                "hospital_mrn": "MRN-77",
                "display_name": "Patient Two",
                "room": "",
                "infectious": False,
                "status": "",
            },
        },
        preview_service=fake_preview,
    )

    request = captured["request"]
    assert request.patient_id == ""
    assert request.patient_type.value == "outpatient"
    assert request.hospital_mrn == "MRN-77"
    assert request.room is None
    assert request.infectious is False
    assert request.status is None
    assert captured["overwrite"] is True
    assert response["ok"] is True
    assert response["kind"] == "patient"
    assert response["action"] == "new_patient"
    assert response["subject_key"] == "98"
    assert response["excel_row"] == 99
    assert response["source_unchanged"] is True
    assert response["verified_in_output"] is True


def test_bridge_keeps_business_validation_out_of_adapter(tmp_path):
    with pytest.raises(RegistrationBridgeError, match="display_name is required"):
        run_registration_bridge(
            {
                "source_path": "baseline.xlsm",
                "preview_dir": str(tmp_path),
                "action": "new_patient",
                "values": {
                    "patient_id": "",
                    "patient_type": "Εσωτερικός",
                    "display_name": "",
                },
            },
            preview_service=lambda *args, **kwargs: pytest.fail("preview should not run"),
        )


@pytest.mark.parametrize(
    ("payload", "message"),
    [
        ({"preview_dir": "x", "action": "new_patient", "values": {}}, "source_path is required"),
        ({"source_path": "x", "action": "new_patient", "values": {}}, "preview_dir is required"),
        ({"source_path": "x", "preview_dir": "y", "values": {}}, "action is required"),
        ({"source_path": "x", "preview_dir": "y", "action": "new_patient", "values": []}, "values must be an object"),
        ({"source_path": "x", "preview_dir": "y", "action": "new_patient", "values": {}, "overwrite": "yes"}, "overwrite must be a boolean"),
    ],
)
def test_bridge_rejects_malformed_ui_payloads(payload, message):
    with pytest.raises(RegistrationBridgeError, match=message):
        run_registration_bridge(payload)
