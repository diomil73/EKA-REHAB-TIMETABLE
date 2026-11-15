"""Fully synthetic ZIP-based XLSM structural preflight regressions."""
from pathlib import Path
from zipfile import ZipFile

import pytest

from rehab_excel.afternoon_native_preflight import (
    AfternoonNativePreflightError, inspect_afternoon_native_candidate,
)


def synthetic_xlsm(tmp_path, *, sheets=None, vba=True, malformed=False):
    if sheets is None:
        sheets = ("PATIENTS", "PATIENT_PLANNER", "MASTER_SCHEDULE")
    output = tmp_path / "SYNTH_NATIVE_CANDIDATE.xlsm"
    xml = ('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
           '<sheets>' + "".join(
               f'<sheet name="{name}" sheetId="{index}"/>'
               for index, name in enumerate(sheets, start=1)
           ) + '</sheets></workbook>')
    with ZipFile(output, "w") as archive:
        archive.writestr("xl/workbook.xml", "<corrupt" if malformed else xml)
        if vba:
            archive.writestr("xl/vbaProject.bin", b"synthetic-not-valid-VBA")
    return output


def test_structurally_eligible_synthetic_candidate_is_read_only(tmp_path):
    path = synthetic_xlsm(tmp_path)
    original = path.read_bytes()
    result = inspect_afternoon_native_candidate(path)
    assert result.ready_for_provenance_review
    assert result.needs_sheet_install
    assert result.has_vba_binary
    assert result.sheet_count == 3
    assert path.read_bytes() == original


def test_existing_afternoon_sheet_blocks_first_time_install(tmp_path):
    path = synthetic_xlsm(
        tmp_path,
        sheets=("PATIENTS", "PATIENT_PLANNER", "MASTER_SCHEDULE", "AFTERNOON_PHYSIO_SLOTS"),
    )
    with pytest.raises(AfternoonNativePreflightError, match="Existing afternoon"):
        inspect_afternoon_native_candidate(path)


@pytest.mark.parametrize("sheets", [
    ("PATIENTS", "MASTER_SCHEDULE"),
    ("PATIENT_PLANNER", "MASTER_SCHEDULE"),
    ("PATIENTS", "PATIENT_PLANNER"),
])
def test_missing_essential_sheet_fails_closed(tmp_path, sheets):
    path = synthetic_xlsm(tmp_path, sheets=sheets)
    with pytest.raises(AfternoonNativePreflightError):
        inspect_afternoon_native_candidate(path)


def test_missing_vba_project_binary_fails_closed(tmp_path):
    path = synthetic_xlsm(tmp_path, vba=False)
    with pytest.raises(AfternoonNativePreflightError):
        inspect_afternoon_native_candidate(path)


def test_duplicate_sheet_roster_fails_closed(tmp_path):
    path = synthetic_xlsm(
        tmp_path,
        sheets=("PATIENTS", "PATIENT_PLANNER", "MASTER_SCHEDULE", "PATIENTS"),
    )
    with pytest.raises(AfternoonNativePreflightError):
        inspect_afternoon_native_candidate(path)


def test_invalid_workbook_xml_fails_closed(tmp_path):
    path = synthetic_xlsm(tmp_path, malformed=True)
    with pytest.raises(AfternoonNativePreflightError):
        inspect_afternoon_native_candidate(path)


def test_refuses_wrong_extension_and_missing_file(tmp_path):
    path = tmp_path / "SYNTH.xlsx"
    path.write_bytes(b"not an XLSM")
    with pytest.raises(AfternoonNativePreflightError):
        inspect_afternoon_native_candidate(path)
    with pytest.raises(AfternoonNativePreflightError):
        inspect_afternoon_native_candidate(tmp_path / "missing.xlsm")


def test_refuses_fake_zip_bytes(tmp_path):
    path = tmp_path / "FAKE.xlsm"
    path.write_bytes(b"not a zip")
    with pytest.raises(AfternoonNativePreflightError):
        inspect_afternoon_native_candidate(path)


def test_no_patient_rows_or_identifiers_returned_in_result(tmp_path):
    path = synthetic_xlsm(tmp_path)
    result = inspect_afternoon_native_candidate(path)
    assert not hasattr(result, "patients")
    assert not hasattr(result, "patient_ids")
