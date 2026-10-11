from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

from rehab_excel import patient_registration_scheduling_slice as scheduling_slice


def test_scheduling_preview_builds_missing_patient_form_on_private_prerequisite(tmp_path, monkeypatch):
    source = tmp_path / "source.xlsm"
    output = tmp_path / "preview.xlsm"
    source.write_bytes(b"baseline")

    observed = {}

    def fake_install_base(path: Path) -> None:
        observed["prerequisite"] = path
        assert path != source
        assert path.read_bytes() == b"baseline"
        path.write_bytes(b"with-patient-form")

    def fake_vertical_preview(source_path, output_path, *, overwrite=False):
        source_path = Path(source_path)
        output_path = Path(output_path)
        observed["vertical_source"] = source_path
        assert source_path.read_bytes() == b"with-patient-form"
        output_path.write_bytes(b"vertical-preview")
        return SimpleNamespace(output_path=str(output_path))

    def fake_install_scheduling(path: Path) -> None:
        observed["scheduling_output"] = path
        assert path == output
        assert path.read_bytes() == b"vertical-preview"

    monkeypatch.setattr(scheduling_slice, "_install_base_patient_form", fake_install_base)
    monkeypatch.setattr(
        scheduling_slice._vertical,
        "create_patient_registration_bridge_preview",
        fake_vertical_preview,
    )
    monkeypatch.setattr(scheduling_slice, "_install_scheduling_ui", fake_install_scheduling)

    report = scheduling_slice.create_patient_registration_bridge_preview(
        source,
        output,
        overwrite=True,
    )

    prerequisite = observed["prerequisite"]
    assert observed["vertical_source"] == prerequisite
    assert observed["scheduling_output"] == output
    assert report.output_path == str(output)
    assert source.read_bytes() == b"baseline"
    assert not prerequisite.exists()
