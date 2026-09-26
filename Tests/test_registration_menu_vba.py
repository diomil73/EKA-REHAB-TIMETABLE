from pathlib import Path
from zipfile import ZipFile

import pytest

from rehab_excel.registration_menu_vba import (
    DEFAULT_PREVIEW_FILENAME,
    MENU_FORM_NAME,
    MENU_MACRO_NAME,
    MENU_MODULE_NAME,
    STANDARD_MODULE_CODE,
    USERFORM_CODE,
    RegistrationMenuVbaError,
    create_registration_menu_preview,
    workbook_has_vba,
)


class FakeInstaller:
    def __init__(self, *, module_present=True, form_present=True):
        self.module_present = module_present
        self.form_present = form_present
        self.paths = []

    def install(self, workbook_path: Path):
        self.paths.append(workbook_path)
        return self.module_present, self.form_present


def _make_fake_xlsm(path: Path, *, with_vba: bool = True) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        if with_vba:
            archive.writestr("xl/vbaProject.bin", b"fake-vba")


def test_menu_component_names_are_stable():
    assert MENU_FORM_NAME == "frmRegistrationMenu"
    assert MENU_MODULE_NAME == "modRegistrationMenu"
    assert MENU_MACRO_NAME == "ShowRegistrationMenu"
    assert DEFAULT_PREVIEW_FILENAME == "REGISTRATION_MENU_PREVIEW.xlsm"


def test_standard_module_exposes_show_macro():
    assert f"Public Sub {MENU_MACRO_NAME}()" in STANDARD_MODULE_CODE
    assert f"{MENU_FORM_NAME}.Show" in STANDARD_MODULE_CODE


def test_userform_code_contains_all_four_click_handlers():
    assert "Private Sub cmdPatient_Click()" in USERFORM_CODE
    assert "Private Sub cmdTherapist_Click()" in USERFORM_CODE
    assert "Private Sub cmdStudent_Click()" in USERFORM_CODE
    assert "Private Sub cmdClose_Click()" in USERFORM_CODE
    assert "Unload Me" in USERFORM_CODE


def test_workbook_has_vba_detects_macro_project(tmp_path):
    with_vba = tmp_path / "with.xlsm"
    without_vba = tmp_path / "without.xlsm"
    _make_fake_xlsm(with_vba, with_vba=True)
    _make_fake_xlsm(without_vba, with_vba=False)

    assert workbook_has_vba(with_vba) is True
    assert workbook_has_vba(without_vba) is False


def test_preview_copies_source_uses_output_only_and_keeps_source_unchanged(tmp_path):
    source = tmp_path / "source.xlsm"
    output = tmp_path / "preview.xlsm"
    _make_fake_xlsm(source, with_vba=True)
    source_before = source.read_bytes()
    backend = FakeInstaller()

    report = create_registration_menu_preview(source, output, backend=backend)

    assert backend.paths == [output.resolve()]
    assert source.read_bytes() == source_before
    assert output.exists()
    assert report.source_unchanged is True
    assert report.vba_present is True
    assert report.module_present is True
    assert report.form_present is True
    assert report.macro_name == MENU_MACRO_NAME


def test_preview_refuses_source_equal_to_output(tmp_path):
    source = tmp_path / "same.xlsm"
    _make_fake_xlsm(source)

    with pytest.raises(RegistrationMenuVbaError, match="different from source"):
        create_registration_menu_preview(source, source, backend=FakeInstaller())


def test_preview_refuses_existing_output_without_overwrite(tmp_path):
    source = tmp_path / "source.xlsm"
    output = tmp_path / "preview.xlsm"
    _make_fake_xlsm(source)
    _make_fake_xlsm(output)

    with pytest.raises(RegistrationMenuVbaError, match="Output already exists"):
        create_registration_menu_preview(source, output, backend=FakeInstaller())


def test_preview_allows_existing_output_with_overwrite(tmp_path):
    source = tmp_path / "source.xlsm"
    output = tmp_path / "preview.xlsm"
    _make_fake_xlsm(source)
    output.write_bytes(b"old")

    report = create_registration_menu_preview(
        source,
        output,
        backend=FakeInstaller(),
        overwrite=True,
    )

    assert report.source_unchanged is True
    assert workbook_has_vba(output) is True


def test_preview_removes_failed_output_if_components_not_confirmed(tmp_path):
    source = tmp_path / "source.xlsm"
    output = tmp_path / "preview.xlsm"
    _make_fake_xlsm(source)

    with pytest.raises(RegistrationMenuVbaError, match="verification failed"):
        create_registration_menu_preview(
            source,
            output,
            backend=FakeInstaller(module_present=False),
        )

    assert not output.exists()


def test_preview_requires_xlsm_extension(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "preview.xlsm"
    source.write_bytes(b"x")

    with pytest.raises(RegistrationMenuVbaError, match="both be .xlsm"):
        create_registration_menu_preview(source, output, backend=FakeInstaller())
