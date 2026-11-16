"""Synthetic-only safety tests for the opt-in private native DEMO shell.

A fake VBA ZIP part is for READ-ONLY structural gate tests; it is not an Excel
macro workbook and is NEVER passed to a real native COM writer.
"""
from datetime import time
from pathlib import Path
from zipfile import ZipFile

import pytest
from openpyxl import Workbook, load_workbook

from rehab_core.programme_mapping import SourceWeeklyEntry
from rehab_excel.trial_demo_native_seed import (
    TrialDemoNativeSeedError, assert_clean_native_demo_shell,
    seed_trial_demo_workbook,
)
from rehab_excel.trial_import_plan import TrialImportPlan
from rehab_excel.weekly_slots_source import HEADERS as WEEKLY_HEADERS


def _blank_shell(tmp_path: Path) -> Path:
    wb = Workbook()
    patients = wb.active
    patients.title = "PATIENTS"
    patients.append(["PatientID", "Θάλαμος", "Ασθενής", "Λοιμώδης", "Κατάσταση"])
    planner = wb.create_sheet("PATIENT_PLANNER")
    planner.append(["PatientID", "Θάλαμος", "Ασθενής", "Λοιμώδης", "Κατάσταση", "ΦΘ_Ώρα"])
    for name in ("MASTER_SCHEDULE", "DAILY_INPUT", "THERAPIST_DAILY", "SETTINGS"):
        wb.create_sheet(name)
    path = tmp_path / "SYNTHETIC_SHELL.xlsm"
    wb.save(path)
    wb.close()
    with ZipFile(path, "a") as zip_file:
        zip_file.writestr("xl/vbaProject.bin", b"DEMO FAKE VBA: STRUCTURE ONLY")
    return path


def _plan(*, shell_eligible=True) -> TrialImportPlan:
    return TrialImportPlan(
        "a" * 64, "b" * 64, 1, 1,
        (SourceWeeklyEntry(
            "DEMO-1", "ΦΘ", "ΔΕ", time(9), "DEMO THERAPIST", "C4",
        ),),
        programme_sha256="c" * 64,
        empty_demo_shell=shell_eligible,
    )


def test_clean_empty_macro_shell_readonly_gate_passes_synthetic_structure(tmp_path):
    path = _blank_shell(tmp_path)
    original = path.read_bytes()
    assert_clean_native_demo_shell(path)
    assert path.read_bytes() == original


def test_nonempty_patient_registry_never_qualifies_as_empty_shell(tmp_path):
    path = _blank_shell(tmp_path)
    wb = load_workbook(path, keep_vba=True)
    wb["PATIENTS"].append(["DEMO-1", "A-1", "DEMO PERSON", "Ο", "Ενεργός"])
    wb.save(path)
    wb.close()
    with pytest.raises(TrialDemoNativeSeedError, match="existing clinical state"):
        assert_clean_native_demo_shell(path)


def test_existing_planner_treatment_blocks_even_if_registry_empty(tmp_path):
    path = _blank_shell(tmp_path)
    wb = load_workbook(path, keep_vba=True)
    wb["PATIENT_PLANNER"]["F2"] = "09:00"
    wb.save(path)
    wb.close()
    with pytest.raises(TrialDemoNativeSeedError, match="old treatments"):
        assert_clean_native_demo_shell(path)


def test_preexisting_weekly_slots_block_instead_of_being_replaced(tmp_path):
    path = _blank_shell(tmp_path)
    wb = load_workbook(path, keep_vba=True)
    ws = wb.create_sheet("PROGRAM_WEEKLY_SLOTS")
    ws.append(WEEKLY_HEADERS)
    ws.append(["DEMO-1", "ΦΘ", "ΔΕ", "09:00", "DEMO T", "C4"])
    wb.save(path)
    wb.close()
    with pytest.raises(TrialDemoNativeSeedError, match="existing weekly slots"):
        assert_clean_native_demo_shell(path)


def test_native_seeder_requires_explicit_empty_shell_plan_without_creating_file(tmp_path):
    out = tmp_path / "NEW_DEMO_PRIVATE.xlsm"
    with pytest.raises(TrialDemoNativeSeedError, match="unsafe demo output"):
        seed_trial_demo_workbook(
            _plan(shell_eligible=False),
            programme=tmp_path / "programme.xlsx",
            review=tmp_path / "review.xlsx",
            intake=tmp_path / "intake.xlsm",
            output=out,
            clinical_approvals={},
        )
    assert not out.exists()


def test_native_seeder_never_overwrites_existing_demo(tmp_path):
    out = tmp_path / "NEW_DEMO_PRIVATE.xlsm"
    out.write_bytes(b"existing unrelated file")
    with pytest.raises(TrialDemoNativeSeedError, match="unsafe demo output"):
        seed_trial_demo_workbook(
            _plan(),
            programme=tmp_path / "programme.xlsx",
            review=tmp_path / "review.xlsx",
            intake=tmp_path / "intake.xlsm",
            output=out,
            clinical_approvals={},
        )
    assert out.read_bytes() == b"existing unrelated file"


def test_native_seeder_rejects_public_git_checkout_path(tmp_path):
    folder = tmp_path / "public-repository"
    folder.mkdir()
    (folder / ".git").mkdir()
    out = folder / "PATIENTS_DEMO_2026.xlsm"
    with pytest.raises(TrialDemoNativeSeedError, match="unsafe demo output"):
        seed_trial_demo_workbook(
            _plan(),
            programme=tmp_path / "programme.xlsx",
            review=tmp_path / "review.xlsx",
            intake=tmp_path / "intake.xlsm",
            output=out,
            clinical_approvals={},
        )
    assert not out.exists()


@pytest.mark.parametrize("sheet", [
    "OUTPATIENT_SCHEDULE", "DAILY_CHANGE_LEDGER", "STUDENTS", "SESSIONS",
    "PATIENT_EXITS",
])
def test_old_operational_sources_are_not_inherited_into_new_demo(tmp_path, sheet):
    path = _blank_shell(tmp_path)
    wb = load_workbook(path, keep_vba=True)
    ws = wb.create_sheet(sheet)
    ws.append(["SampleHeader"])
    ws.append(["OLD SYNTHETIC RECORD"])
    wb.save(path)
    wb.close()
    with pytest.raises(TrialDemoNativeSeedError, match="old operational records"):
        assert_clean_native_demo_shell(path)


def test_preexisting_patient_exit_events_block_new_demo_shell(tmp_path):
    from rehab_excel.patient_exit_source import HEADERS as PATIENT_EXIT_HEADERS
    path = _blank_shell(tmp_path)
    wb = load_workbook(path, keep_vba=True)
    sheet = wb.create_sheet("PATIENT_EXITS")
    sheet.append(PATIENT_EXIT_HEADERS)
    sheet.append(["DEMO-PATIENT", "inpatient_discharge", "2026-11-16", "DEMO"])
    wb.save(path)
    wb.close()
    with pytest.raises(TrialDemoNativeSeedError, match="old operational records"):
        assert_clean_native_demo_shell(path)


def test_private_native_source_contains_explicit_outpatient_write_and_readback_contract():
    import inspect
    import rehab_excel.trial_demo_native_seed as seed
    text = inspect.getsource(seed.Win32ComTrialDemoSeedBackend.write)
    verify = inspect.getsource(seed._verify_seed_workbook)
    assert 'Worksheets("OUTPATIENT_SCHEDULE")' in text
    assert "projection.outpatient.appointment_rows" in text
    assert "Native weekly view differs from approved sources" in verify
    assert "Outpatient source readback differs" in verify
    assert "read_unified_base_schedule(stage)" in verify
