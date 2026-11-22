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


def _with_pm_sheet(path, *, state="veryHidden", header=None, extra=None):
    from rehab_excel.afternoon_slots_source import SHEET_NAME, HEADERS
    wb = load_workbook(path, keep_vba=True)
    sheet = wb.create_sheet(SHEET_NAME)
    sheet.append(HEADERS if header is None else header)
    sheet.sheet_state = state
    if extra is not None:
        address, value = extra
        sheet[address] = value
    wb.save(path)
    wb.close()


def test_clean_shell_accepts_existing_empty_canonical_afternoon_source(tmp_path):
    path = _blank_shell(tmp_path)
    _with_pm_sheet(path)
    before = path.read_bytes()
    assert_clean_native_demo_shell(path)
    assert before == path.read_bytes()


@pytest.mark.parametrize("extra", [
    ("A2", "FICTIONAL OLD PATIENT"),
    ("D18", "=TIME(17,30,0)"),
    ("G20", "FICTIONAL EXTRA STATE"),
    ("C7", "#REF!"),
])
def test_old_afternoon_state_blocks_clean_shell_even_outside_first_row(tmp_path, extra):
    path = _blank_shell(tmp_path)
    _with_pm_sheet(path, extra=extra)
    before = path.read_bytes()
    with pytest.raises(TrialDemoNativeSeedError, match="old afternoon records"):
        assert_clean_native_demo_shell(path)
    assert before == path.read_bytes()


@pytest.mark.parametrize("kwargs", [
    {"state": "visible"}, {"header": ["Wrong"]}, {"extra": ("G1", "UNKNOWN")},
])
def test_existing_invalid_afternoon_schema_is_not_repaired(tmp_path, kwargs):
    path = _blank_shell(tmp_path)
    _with_pm_sheet(path, **kwargs)
    with pytest.raises(TrialDemoNativeSeedError, match="schema/visibility"):
        assert_clean_native_demo_shell(path)


def test_readback_requires_independent_empty_afternoon_source():
    from rehab_excel.trial_demo_native_seed import _assert_empty_afternoon_source
    from rehab_excel.afternoon_slots_source import SHEET_NAME, HEADERS, read_afternoon_slots_sheet
    book = Workbook()
    with pytest.raises(TrialDemoNativeSeedError, match="source missing"):
        _assert_empty_afternoon_source(book, required=True)
    sheet = book.create_sheet(SHEET_NAME)
    sheet.append(HEADERS)
    sheet.sheet_state = "veryHidden"
    _assert_empty_afternoon_source(book, required=True)
    assert read_afternoon_slots_sheet(sheet, known_patient_ids={"FICTIONAL-1"}) == ()
    sheet["E8"] = "FICTIONAL PROVIDER"
    with pytest.raises(TrialDemoNativeSeedError, match="old afternoon records"):
        _assert_empty_afternoon_source(book, required=True)
    book.close()


def test_native_seed_wires_empty_pm_initialization_and_readback():
    import inspect
    import rehab_excel.trial_demo_native_seed as seed
    assert "_initialize_empty_afternoon_source(wb)" in inspect.getsource(seed.Win32ComTrialDemoSeedBackend.write)
    assert "_assert_empty_afternoon_source(wb, required=True)" in inspect.getsource(seed._verify_seed_workbook)
    assert "sheet.Visible = 2" in inspect.getsource(seed._initialize_empty_afternoon_source)


def _fictional_full_provider_schedule_snapshot(
    *, shared_members=(), outpatient_conflict=False, same_patient_combined=False,
):
    """All records are fabricated. Source-bound clinical approvals stay explicit."""
    from rehab_excel.trial_provider_schedule import TrialProviderSchedule
    from rehab_excel.trial_demo_source_projection import TrialDemoSourceProjection
    from rehab_excel.trial_patient_registry_projection import (
        PATIENTS_HEADERS, TrialPatientRegistryProjection,
    )
    from rehab_excel.trial_private_outpatient_projection import (
        TrialPrivateOutpatientProjection,
    )
    from rehab_excel.trial_weekly_projection import TrialWeeklyProjection

    def row(pid, patient_type):
        return (pid, None, "SYNTHETIC PERSON", "Ο", None,
                patient_type, None, None, "FICTIONAL")

    inpatient_patients = (
        row("DEMO-A", "Εσωτερικός"), row("DEMO-B", "Εσωτερικός"),
    )
    outpatient_patients = (row("DEMO-O", "Εξωτερικός"),)
    clock = time(9, 30)
    in_rows = [
        ("DEMO-A", "ΦΘ", "ΔΕ", clock, "DEMO STAFF", "C4"),
        ("DEMO-B", "ΦΘ", "ΔΕ", clock, "Demo.Staff", "C5"),
    ]
    if same_patient_combined:
        in_rows.append(
            ("DEMO-A", "Ρομποτικό", "ΔΕ", clock, "DEMO STAFF", "R4")
        )
    out_day = "ΔΕ" if outpatient_conflict else "ΤΡ"
    out_time = clock if outpatient_conflict else time(10, 0)
    outpatient_weekly = (
        ("DEMO-O", "SYNTHETIC PERSON", "Εργο", out_day,
         out_time, "DEMO STAFF", "OUT:2"),
    )
    inpatient = TrialDemoSourceProjection(
        TrialPatientRegistryProjection(
            PATIENTS_HEADERS, inpatient_patients, len(inpatient_patients),
        ),
        TrialWeeklyProjection(
            ("PatientID", "Θεραπεία", "Ημέρα", "Ώρα",
             "Θεραπευτής", "Πηγή"),
            tuple(in_rows), len(in_rows),
        ),
        len(inpatient_patients), len(in_rows),
    )
    outpatient = TrialPrivateOutpatientProjection(
        "d" * 64, outpatient_patients,
        (("DEMO-O", "SYNTHETIC PERSON", "Εργο",
          out_time, out_day, "DEMO STAFF"),),
        outpatient_weekly,
        1, 1, 1,
    )
    source_plan = TrialImportPlan(
        "a" * 64, "b" * 64, 2, len(in_rows), (),
        programme_sha256="c" * 64,
    )
    approved_shared = tuple(
        ("DEMO STAFF", "ΔΕ", clock, frozenset(members))
        for members in shared_members
    )
    schedule = TrialProviderSchedule(
        ("c" * 64, "a" * 64, "b" * 64, "d" * 64),
        (("DEMO STAFF", "ΔΕ", frozenset({clock})),
         ("DEMO STAFF", "ΤΡ", frozenset({time(10, 0)}))),
        approved_shared,
    )
    return inpatient, outpatient, source_plan, schedule


def test_private_provider_schedule_requires_all_inpatient_shared_member_approvals():
    from rehab_excel.trial_provider_schedule import (
        TrialProviderScheduleError, merge_with_approved_provider_schedule,
    )
    inpatient, outpatient, plan, absent = (
        _fictional_full_provider_schedule_snapshot()
    )
    with pytest.raises(TrialProviderScheduleError, match="exact approved membership"):
        merge_with_approved_provider_schedule(
            inpatient, outpatient, plan=plan, schedule=absent,
        )
    inp, out, approved_plan, approved = (
        _fictional_full_provider_schedule_snapshot(
            shared_members=(("DEMO-A", "DEMO-B"),),
        )
    )
    combined = merge_with_approved_provider_schedule(
        inp, out, plan=approved_plan, schedule=approved,
    )
    assert combined.patients == 3
    assert combined.weekday_slot_count == 3
    assert combined.outpatient is out


def test_private_provider_schedule_requires_exact_not_stale_membership():
    from rehab_excel.trial_provider_schedule import (
        TrialProviderScheduleError, merge_with_approved_provider_schedule,
    )
    for members in (
        (("DEMO-A", "DEMO-O"),),
        (("DEMO-A", "DEMO-B", "DEMO-O"),),
    ):
        inp, out, plan, schedule = _fictional_full_provider_schedule_snapshot(
            shared_members=members,
        )
        with pytest.raises(TrialProviderScheduleError, match="exact approved membership"):
            merge_with_approved_provider_schedule(
                inp, out, plan=plan, schedule=schedule,
            )


def test_private_provider_schedule_combined_pair_same_patient_counts_once():
    from rehab_excel.trial_provider_schedule import (
        merge_with_approved_provider_schedule,
    )
    inp, out, plan, schedule = _fictional_full_provider_schedule_snapshot(
        shared_members=(("DEMO-A", "DEMO-B"),),
        same_patient_combined=True,
    )
    combined = merge_with_approved_provider_schedule(
        inp, out, plan=plan, schedule=schedule,
    )
    assert combined.weekday_slot_count == 4


def test_private_provider_schedule_preserves_inpatient_outpatient_shared_gate():
    from rehab_excel.trial_provider_schedule import (
        TrialProviderScheduleError, merge_with_approved_provider_schedule,
    )
    inp, out, plan, missing_outpatient = (
        _fictional_full_provider_schedule_snapshot(
            shared_members=(("DEMO-A", "DEMO-B"),),
            outpatient_conflict=True,
        )
    )
    with pytest.raises(TrialProviderScheduleError, match="exact approved membership"):
        merge_with_approved_provider_schedule(
            inp, out, plan=plan, schedule=missing_outpatient,
        )
    # Source schedule permits at most one approval per provider/day/time.
    # The precise three-patient group is required rather than a second entry.
    inp, out, plan, approved = _fictional_full_provider_schedule_snapshot(
        shared_members=(("DEMO-A", "DEMO-B", "DEMO-O"),),
        outpatient_conflict=True,
    )
    merged = merge_with_approved_provider_schedule(
        inp, out, plan=plan, schedule=approved,
    )
    assert merged.patients == 3
    assert merged.weekday_slot_count == 3
