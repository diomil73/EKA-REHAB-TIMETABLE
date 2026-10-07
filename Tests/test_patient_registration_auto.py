from pathlib import Path

from rehab_core.models import Patient
from rehab_core.registration import NewPatientRequest
from rehab_excel.patient_registration import (
    PatientRegistrationPreviewReport,
    PatientRegistrationWriteError,
)
from rehab_excel.patient_registration_auto import (
    create_auto_patient_registration_preview,
    resolve_patient_id_request,
)
from rehab_excel.master_projection_writer import MasterProjectionWriteError
from rehab_excel.master_visual_style import MasterVisualStyleError
from rehab_excel.patient_planner_projection import PatientPlannerProjectionError


def test_blank_patient_id_is_allocated_sequentially():
    request = NewPatientRequest(patient_id="", display_name="ΝΕΟΣ")
    resolved = resolve_patient_id_request(
        request,
        existing_patients=[Patient("94", "A"), Patient("TEST-X", "B")],
    )
    assert resolved.patient_id == "95"
    assert request.patient_id == ""


def test_explicit_legacy_patient_id_is_preserved():
    request = NewPatientRequest(patient_id="IMPORT-001", display_name="ΝΕΟΣ")
    resolved = resolve_patient_id_request(
        request,
        existing_patients=[Patient("94", "A")],
    )
    assert resolved is request
    assert resolved.patient_id == "IMPORT-001"


def test_auto_patient_preview_refreshes_master_after_registration(monkeypatch, tmp_path):
    source = tmp_path / "source.xlsm"
    output = tmp_path / "preview.xlsm"
    source.write_bytes(b"source")

    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.read_patient_registry",
        lambda _path: [Patient("7", "OLD")],
    )

    captured = {}

    def fake_create(source_path, output_path, request, *, backend, overwrite):
        captured["request"] = request
        Path(output_path).write_bytes(b"preview")
        return PatientRegistrationPreviewReport(
            source_path=str(source_path),
            output_path=str(output_path),
            patient_id=request.patient_id,
            excel_row=8,
            source_unchanged=True,
            verified_in_output=True,
        )

    refreshed = []

    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.create_patient_registration_preview",
        fake_create,
    )
    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.refresh_patient_planner_projection_in_place",
        lambda path: refreshed.append(("planner", Path(path))) or 8,
    )
    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.refresh_master_projection_in_place",
        lambda path: refreshed.append(("master", Path(path))) or 8,
    )
    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.apply_master_visual_style_in_place",
        lambda path: refreshed.append(("style", Path(path))),
    )

    report = create_auto_patient_registration_preview(
        source,
        output,
        NewPatientRequest(patient_id="", display_name="ΝΕΟΣ"),
        overwrite=True,
    )

    assert report.patient_id == "8"
    assert captured["request"].patient_id == "8"
    assert refreshed == [
        ("planner", output),
        ("master", output),
        ("style", output),
    ]


def test_auto_patient_preview_removes_output_if_master_refresh_fails(monkeypatch, tmp_path):
    source = tmp_path / "source.xlsm"
    output = tmp_path / "preview.xlsm"
    source.write_bytes(b"source")

    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.read_patient_registry",
        lambda _path: [],
    )

    def fake_create(source_path, output_path, request, *, backend, overwrite):
        Path(output_path).write_bytes(b"preview")
        return PatientRegistrationPreviewReport(
            source_path=str(source_path),
            output_path=str(output_path),
            patient_id=request.patient_id,
            excel_row=2,
            source_unchanged=True,
            verified_in_output=True,
        )

    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.create_patient_registration_preview",
        fake_create,
    )
    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.refresh_patient_planner_projection_in_place",
        lambda _path: 1,
    )
    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.refresh_master_projection_in_place",
        lambda _path: (_ for _ in ()).throw(MasterProjectionWriteError("boom")),
    )
    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.apply_master_visual_style_in_place",
        lambda _path: None,
    )

    try:
        create_auto_patient_registration_preview(
            source,
            output,
            NewPatientRequest(patient_id="", display_name="ΝΕΟΣ"),
            overwrite=True,
        )
    except PatientRegistrationWriteError as exc:
        assert "projection refresh failed" in str(exc)
    else:
        raise AssertionError("Expected MASTER refresh failure")

    assert not output.exists()


def test_auto_patient_preview_removes_output_if_planner_refresh_fails(monkeypatch, tmp_path):
    source = tmp_path / "source.xlsm"
    output = tmp_path / "preview.xlsm"
    source.write_bytes(b"source")

    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.read_patient_registry",
        lambda _path: [],
    )

    def fake_create(source_path, output_path, request, *, backend, overwrite):
        Path(output_path).write_bytes(b"preview")
        return PatientRegistrationPreviewReport(
            source_path=str(source_path),
            output_path=str(output_path),
            patient_id=request.patient_id,
            excel_row=2,
            source_unchanged=True,
            verified_in_output=True,
        )

    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.create_patient_registration_preview",
        fake_create,
    )
    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.refresh_patient_planner_projection_in_place",
        lambda _path: (_ for _ in ()).throw(PatientPlannerProjectionError("boom")),
    )
    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.refresh_master_projection_in_place",
        lambda _path: 1,
    )
    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.apply_master_visual_style_in_place",
        lambda _path: None,
    )

    try:
        create_auto_patient_registration_preview(
            source,
            output,
            NewPatientRequest(patient_id="", display_name="ΝΕΟΣ"),
            overwrite=True,
        )
    except PatientRegistrationWriteError as exc:
        assert "projection refresh failed" in str(exc)
    else:
        raise AssertionError("Expected planner refresh failure")

    assert not output.exists()


def test_auto_patient_preview_removes_output_if_visual_style_fails(monkeypatch, tmp_path):
    source = tmp_path / "source.xlsm"
    output = tmp_path / "preview.xlsm"
    source.write_bytes(b"source")

    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.read_patient_registry",
        lambda _path: [],
    )

    def fake_create(source_path, output_path, request, *, backend, overwrite):
        Path(output_path).write_bytes(b"preview")
        return PatientRegistrationPreviewReport(
            source_path=str(source_path),
            output_path=str(output_path),
            patient_id=request.patient_id,
            excel_row=2,
            source_unchanged=True,
            verified_in_output=True,
        )

    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.create_patient_registration_preview",
        fake_create,
    )
    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.refresh_patient_planner_projection_in_place",
        lambda _path: 1,
    )
    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.refresh_master_projection_in_place",
        lambda _path: 1,
    )
    monkeypatch.setattr(
        "rehab_excel.patient_registration_auto.apply_master_visual_style_in_place",
        lambda _path: (_ for _ in ()).throw(MasterVisualStyleError("boom")),
    )

    try:
        create_auto_patient_registration_preview(
            source,
            output,
            NewPatientRequest(patient_id="", display_name="ΝΕΟΣ"),
            overwrite=True,
        )
    except PatientRegistrationWriteError as exc:
        assert "projection refresh failed" in str(exc)
    else:
        raise AssertionError("Expected visual style failure")

    assert not output.exists()
