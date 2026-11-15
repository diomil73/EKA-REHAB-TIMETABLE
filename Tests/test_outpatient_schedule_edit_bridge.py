from pathlib import Path

from rehab_excel.outpatient_schedule_bridge import run_outpatient_schedule_bridge
from rehab_excel.outpatient_schedule_registration import OutpatientSchedulePreviewReport


def test_bridge_forwards_target_base_entry_id_for_update(monkeypatch, tmp_path):
    captured = {}

    def fake_preview(
        source,
        output,
        request,
        *,
        overwrite,
        allow_therapist_double_booking,
    ):
        captured["request"] = request
        return OutpatientSchedulePreviewReport(
            source_path=str(source),
            output_path=str(output),
            base_entry_id="outpatient:4",
            excel_row=4,
            source_unchanged=True,
            verified_in_output=True,
            created=False,
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
                "patient_id": "P-OUT",
                "treatment": "ΦΘ",
                "time": "10:30",
                "days": "Τρ-Πε",
                "therapist": "T2",
                "target_base_entry_id": "outpatient:4",
            },
        },
        fingerprint_service=lambda _: "abc123",
    )

    request = captured["request"]
    assert request.target_base_entry_id == "outpatient:4"
    assert response["created"] is False
    assert response["base_entry_id"] == "outpatient:4"
    assert Path(response["output_path"]).name == "OUTPATIENT_SCHEDULE_PREVIEW.xlsm"
