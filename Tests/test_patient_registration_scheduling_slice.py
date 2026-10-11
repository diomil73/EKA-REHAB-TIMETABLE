from __future__ import annotations

import inspect
from pathlib import Path
from types import SimpleNamespace

import pytest

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
        path.write_bytes(b"with-scheduling-ui")

    def fake_refresh_daily(path: Path) -> None:
        observed["daily_refresh_output"] = path
        assert path == output
        assert path.read_bytes() == b"with-scheduling-ui"
        path.write_bytes(b"with-fresh-daily-visuals")

    monkeypatch.setattr(scheduling_slice, "_install_base_patient_form", fake_install_base)
    monkeypatch.setattr(
        scheduling_slice._vertical,
        "create_patient_registration_bridge_preview",
        fake_vertical_preview,
    )
    monkeypatch.setattr(scheduling_slice, "_install_scheduling_ui", fake_install_scheduling)
    monkeypatch.setattr(scheduling_slice, "_refresh_daily_visuals", fake_refresh_daily)

    report = scheduling_slice.create_patient_registration_bridge_preview(
        source,
        output,
        overwrite=True,
    )

    prerequisite = observed["prerequisite"]
    assert observed["vertical_source"] == prerequisite
    assert observed["scheduling_output"] == output
    assert observed["daily_refresh_output"] == output
    assert output.read_bytes() == b"with-fresh-daily-visuals"
    assert report.output_path == str(output)
    assert source.read_bytes() == b"baseline"
    assert not prerequisite.exists()


def test_scheduling_ui_installer_refreshes_patient_edit_form_and_installs_master_renderer():
    source = inspect.getsource(scheduling_slice._install_scheduling_ui)
    assert "install_patient_edit_form" in source
    assert "install_patient_program_master_render(vbproject)" in source
    assert 'Columns("D:J").NumberFormat = "hh:mm"' not in source
    assert 'stage = "installing MASTER patient-program renderer"' in source
    assert 'stage = "verifying patient-program VBA replacement"' in source
    assert "_verify_patient_program_vba(vbproject)" in source


def test_legacy_patient_scheduling_runtime_is_removed_before_new_runtime_install():
    legacy = object()
    remaining = {
        scheduling_slice.LEGACY_PATIENT_SCHEDULING_MODULE_NAME: legacy,
        "PatientProgramRuntime": object(),
    }
    removed = []

    class FakeVBComponents:
        def __call__(self, name):
            if name not in remaining:
                raise KeyError(name)
            return remaining[name]

        def Remove(self, component):
            removed.append(component)
            for name, value in tuple(remaining.items()):
                if value is component:
                    del remaining[name]
                    return
            raise AssertionError("component was not present")

    vbproject = SimpleNamespace(VBComponents=FakeVBComponents())

    assert scheduling_slice._remove_vb_component_if_present(
        vbproject,
        scheduling_slice.LEGACY_PATIENT_SCHEDULING_MODULE_NAME,
    ) is True
    assert removed == [legacy]
    assert scheduling_slice.LEGACY_PATIENT_SCHEDULING_MODULE_NAME not in remaining
    assert "PatientProgramRuntime" in remaining

    assert scheduling_slice._remove_vb_component_if_present(
        vbproject,
        scheduling_slice.LEGACY_PATIENT_SCHEDULING_MODULE_NAME,
    ) is False


def _fake_component(code: str):
    class FakeCodeModule:
        CountOfLines = 1

        def Lines(self, start, count):
            assert start == 1
            assert count == 1
            return code

    return SimpleNamespace(CodeModule=FakeCodeModule())


def test_patient_program_vba_verification_accepts_only_new_runtime_and_form():
    components = {
        scheduling_slice.PATIENT_PROGRAM_FORM_NAME: _fake_component(
            "lstTreatments\nSavePatientProgramAtomic\nBuildAutomaticRequestArguments"
        ),
        scheduling_slice.PATIENT_PROGRAM_MODULE_NAME: _fake_component(
            "LoadPatientProgramConfig\nCompletePatientProgram\nOpenPatientScheduling"
        ),
    }

    class FakeComponents:
        def __call__(self, name):
            if name not in components:
                raise KeyError(name)
            return components[name]

    scheduling_slice._verify_patient_program_vba(
        SimpleNamespace(VBComponents=FakeComponents())
    )


def test_patient_program_vba_verification_rejects_legacy_form_code():
    components = {
        scheduling_slice.PATIENT_PROGRAM_FORM_NAME: _fake_component(
            "lstTreatments\nSavePatientProgramAtomic\nBuildAutomaticRequestArguments\n"
            "cboTreatment\nLoadRecurringSchedulingOptions"
        ),
        scheduling_slice.PATIENT_PROGRAM_MODULE_NAME: _fake_component(
            "LoadPatientProgramConfig\nCompletePatientProgram\nOpenPatientScheduling"
        ),
    }

    class FakeComponents:
        def __call__(self, name):
            if name not in components:
                raise KeyError(name)
            return components[name]

    with pytest.raises(RuntimeError, match="legacy markers still present"):
        scheduling_slice._verify_patient_program_vba(
            SimpleNamespace(VBComponents=FakeComponents())
        )


def test_recurring_visual_refresh_uses_unified_inpatient_and_outpatient_schedule(tmp_path, monkeypatch):
    workbook = tmp_path / "preview.xlsm"
    workbook.write_bytes(b"preview")

    patients = [object()]
    unified_entries = ["planner-entry", "outpatient-entry"]
    observed = {}

    monkeypatch.setattr(
        scheduling_slice,
        "read_patient_registry",
        lambda path: patients,
    )
    monkeypatch.setattr(
        scheduling_slice,
        "read_unified_base_schedule",
        lambda path: unified_entries,
    )

    def fake_style_plan(path, entries, patient_list):
        observed["path"] = Path(path)
        observed["entries"] = entries
        observed["patients"] = patient_list
        return SimpleNamespace(operation_count=0)

    monkeypatch.setattr(
        scheduling_slice,
        "build_recurring_therapist_daily_style_plan",
        fake_style_plan,
    )

    scheduling_slice._refresh_daily_visuals(workbook)

    assert observed["path"] == workbook
    assert observed["entries"] is unified_entries
    assert observed["patients"] is patients
