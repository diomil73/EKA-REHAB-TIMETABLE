"""Synthetic-only read-only patient registry projection acceptance cases."""
from datetime import time
from hashlib import sha256

import pytest
from openpyxl import Workbook, load_workbook

from rehab_core.models import PatientType
from rehab_core.programme_mapping import SourceWeeklyEntry
from rehab_excel.trial_import_plan import TrialImportPlan
from rehab_excel.trial_patient_registry_projection import (
    ConfirmedClinicalFields,
    PATIENTS_HEADERS,
    TrialPatientRegistryProjectionError,
    project_trial_patient_registry,
)
from rehab_excel.trial_source_alignment import (
    PATIENT_HEADERS, PROGRAMME_HEADERS, WEEKLY_HEADERS,
)
from rehab_excel.patient_registry_source import read_patient_registry


def _fixture(tmp_path, *, source_context="UNIT A / DR DEMO", confirm="CONFIRMED"):
    programme = tmp_path / "programme.xlsx"
    review = tmp_path / "review.xlsx"
    intake = tmp_path / "intake.xlsm"

    original = Workbook()
    src = original.active
    src.title = "ΠΡΟΓΡΑΜΜΑ"
    for column, value in enumerate(PROGRAMME_HEADERS, start=3):
        src.cell(1, column, value)
    src["A4"] = source_context
    src["B4"] = "DEMO PERSON"
    src["C4"] = "09:30"
    src["D4"] = "ΔΕ"
    src["E4"] = "DEMO THERAPIST"
    original.save(programme)
    original.close()

    staged = Workbook()
    patients = staged.active
    patients.title = "ΑΣΘΕΝΕΙΣ"
    patients.append(PATIENT_HEADERS)
    patients.append([4, "CLINIC-A", "DEMO PERSON", source_context, "DEMO-1", confirm])
    weekly = staged.create_sheet("ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ")
    weekly.append(WEEKLY_HEADERS)
    weekly.append([
        "CLINIC-A", 4, "DEMO PERSON", "ΦΘ", "ΔΕ", "09:30",
        "DEMO THERAPIST", None, "C4 / D4 / E4", "Synthetic", "PROPOSED_EXACT_SOURCE",
    ])
    staged.save(review)
    staged.close()
    intake.write_bytes(b"synthetic-intake-do-not-open")
    digest = lambda path: sha256(path.read_bytes()).hexdigest()
    plan = TrialImportPlan(
        review_sha256=digest(review), intake_sha256=digest(intake),
        patient_count=1, weekday_slot_count=1,
        entries=(SourceWeeklyEntry(
            "DEMO-1", "ΦΘ", "ΔΕ", time(9, 30), "DEMO THERAPIST", "C4"
        ),),
        programme_sha256=digest(programme),
    )
    return plan, programme, review, intake


def _approval(**overrides):
    attrs = dict(
        patient_type=PatientType.INPATIENT, infectious=False,
        room=None, responsible_doctor=None, status=None,
    )
    attrs.update(overrides)
    return {("CLINIC-A", 4): ConfirmedClinicalFields(**attrs)}


def _run(fixture, *, approval=None):
    plan, programme, review, intake = fixture
    return project_trial_patient_registry(
        plan, programme=programme, review=review, intake=intake,
        clinical_approvals=_approval() if approval is None else approval,
    )


def test_synthetic_registry_projection_matches_existing_patient_reader(tmp_path):
    fixture = _fixture(tmp_path)
    paths = fixture[1:]
    before = tuple(path.read_bytes() for path in paths)
    result = _run(fixture, approval=_approval(
        room="A-101", responsible_doctor="DEMO DOCTOR", infectious=True
    ))
    assert result.headers == PATIENTS_HEADERS
    assert result.patient_count == 1
    assert result.rows[0] == (
        "DEMO-1", "A-101", "DEMO PERSON", "Ν", None,
        "Εσωτερικός", None, "DEMO DOCTOR", "UNIT A / DR DEMO"
    )
    assert "DEMO PERSON" not in repr(result)
    assert "DEMO-1" not in repr(result)
    assert tuple(path.read_bytes() for path in paths) == before
    # In-memory roundtrip to the existing PATIENTS reader on a disposable
    # synthetic XLSX, never through a native patient import.
    wb = Workbook()
    ws = wb.active
    ws.title = "PATIENTS"
    ws.append(result.headers)
    for row in result.rows:
        ws.append(row)
    disposable = tmp_path / "synthetic-projected-patients.xlsx"
    wb.save(disposable)
    wb.close()
    [patient] = read_patient_registry(disposable)
    assert patient.patient_type is PatientType.INPATIENT
    assert patient.infectious
    assert patient.room == "A-101"
    assert patient.responsible_doctor == "DEMO DOCTOR"


def test_source_context_blank_is_not_forward_filled_or_guessed(tmp_path):
    fixture = _fixture(tmp_path, source_context=None)
    result = _run(fixture)
    assert result.rows[0][1] is None
    assert result.rows[0][7] is None
    assert result.rows[0][8] is None


def test_unconfirmed_identity_never_enters_registry(tmp_path):
    fixture = _fixture(tmp_path, confirm="ΕΚΚΡΕΜΕΙ")
    with pytest.raises(TrialPatientRegistryProjectionError, match="Unconfirmed"):
        _run(fixture)


def test_missing_or_surplus_clinical_approval_blocks_whole_registry(tmp_path):
    fixture = _fixture(tmp_path)
    with pytest.raises(TrialPatientRegistryProjectionError, match="classification"):
        _run(fixture, approval={})
    extra = _approval()
    extra[("CLINIC-B", 5)] = ConfirmedClinicalFields(
        patient_type=PatientType.INPATIENT, infectious=False
    )
    with pytest.raises(TrialPatientRegistryProjectionError, match="scope"):
        _run(fixture, approval=extra)


def test_source_hash_change_and_missing_bound_source_stop(tmp_path):
    fixture = _fixture(tmp_path)
    plan, programme, review, intake = fixture
    with pytest.raises(TrialPatientRegistryProjectionError, match="Source-bound"):
        _run((TrialImportPlan(
            plan.review_sha256, plan.intake_sha256, 1, 1, plan.entries
        ), programme, review, intake))
    programme.write_bytes(programme.read_bytes() + b"unexpected-change")
    with pytest.raises(TrialPatientRegistryProjectionError, match="SAFETY STOP"):
        _run(fixture)


def test_outpatient_must_not_have_infectious_or_room(tmp_path):
    fixture = _fixture(tmp_path)
    with pytest.raises(TrialPatientRegistryProjectionError, match="contradictory"):
        _run(fixture, approval=_approval(
            patient_type=PatientType.OUTPATIENT, infectious=True,
        ))
    with pytest.raises(TrialPatientRegistryProjectionError, match="contradictory"):
        _run(fixture, approval=_approval(
            patient_type=PatientType.OUTPATIENT, room="A-101",
        ))
    result = _run(fixture, approval=_approval(patient_type=PatientType.OUTPATIENT))
    assert result.rows[0][5] == "Εξωτερικός"
    assert result.rows[0][1] is None


def test_unverified_source_change_or_review_mapping_blocks(tmp_path):
    fixture = _fixture(tmp_path)
    plan, programme, review, intake = fixture
    wb = load_workbook(review)
    wb["ΑΣΘΕΝΕΙΣ"]["D2"] = "INCORRECT CONTEXT"
    wb.save(review)
    wb.close()
    with pytest.raises(TrialPatientRegistryProjectionError, match="SAFETY STOP"):
        _run(fixture)


def test_patient_registry_plan_must_cover_all_admitted_weekly_ids(tmp_path):
    fixture = _fixture(tmp_path)
    plan, programme, review, intake = fixture
    bad = TrialImportPlan(
        review_sha256=plan.review_sha256,
        intake_sha256=plan.intake_sha256,
        patient_count=1,
        weekday_slot_count=1,
        entries=(SourceWeeklyEntry(
            "UNCONFIRMED", "ΦΘ", "ΔΕ", time(9, 30),
            "DEMO THERAPIST", "C4",
        ),),
        programme_sha256=plan.programme_sha256,
    )
    with pytest.raises(TrialPatientRegistryProjectionError, match="scope"):
        _run((bad, programme, review, intake))


def test_joined_demo_projection_preserves_both_sheets_and_roundtrips(tmp_path):
    from rehab_excel.trial_demo_source_projection import project_trial_demo_sources
    from rehab_excel.weekly_slots_source import read_weekly_slots_sheet
    plan, programme, review, intake = _fixture(tmp_path)
    joined = project_trial_demo_sources(
        plan, programme=programme, review=review, intake=intake,
        clinical_approvals=_approval(),
    )
    assert joined.patients == joined.patient_registry.patient_count == 1
    assert joined.weekday_slot_count == joined.weekly_slots.row_count == 1
    assert "DEMO-1" not in repr(joined)
    assert joined.patient_registry.rows[0][0] == joined.weekly_slots.rows[0][0]
    weekly = Workbook()
    ws = weekly.active
    ws.title = "PROGRAM_WEEKLY_SLOTS"
    ws.append(joined.weekly_slots.headers)
    for row in joined.weekly_slots.rows:
        ws.append(row)
    result = read_weekly_slots_sheet(ws, inpatient_ids={"DEMO-1"})
    assert len(result) == 1
    assert (result[0].day_pattern, result[0].start_time) == ("Δε", time(9, 30))
    weekly.close()


def test_joined_demo_projection_never_silently_imports_outpatient_slots(tmp_path):
    from rehab_excel.trial_demo_source_projection import project_trial_demo_sources
    from rehab_excel.trial_weekly_projection import TrialWeeklyProjectionError
    plan, programme, review, intake = _fixture(tmp_path)
    with pytest.raises(TrialWeeklyProjectionError, match="inpatient"):
        project_trial_demo_sources(
            plan, programme=programme, review=review, intake=intake,
            clinical_approvals=_approval(patient_type=PatientType.OUTPATIENT),
        )


def test_source_parity_requires_all_five_days_when_source_day_is_blank(tmp_path):
    from rehab_excel.trial_demo_source_reconciliation import (
        TrialDemoSourceReconciliationError, reconcile_demo_source_weekdays,
    )
    from rehab_excel.trial_demo_source_projection import project_trial_demo_sources
    plan, programme, review, intake = _fixture(tmp_path)
    wb = load_workbook(programme)
    wb["ΠΡΟΓΡΑΜΜΑ"]["D4"] = None
    wb.save(programme)
    wb.close()
    plan = TrialImportPlan(
        plan.review_sha256, plan.intake_sha256, 1, 1, plan.entries,
        sha256(programme.read_bytes()).hexdigest(),
    )
    with pytest.raises(TrialDemoSourceReconciliationError, match="projection differs"):
        project_trial_demo_sources(
            plan, programme=programme, review=review, intake=intake,
            clinical_approvals=_approval(),
        )
    staged = load_workbook(review)
    ws = staged["ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"]
    for day in ("ΤΡ", "ΤΕ", "ΠΕ", "ΠΑ"):
        ws.append([
            "CLINIC-A", 4, "DEMO PERSON", "ΦΘ", day, "09:30",
            "DEMO THERAPIST", None, "C4 / D4 / E4",
            "Synthetic", "PROPOSED_EXACT_SOURCE",
        ])
    staged.save(review)
    staged.close()
    entries = tuple(
        SourceWeeklyEntry(
            "DEMO-1", "ΦΘ", day, time(9, 30), "DEMO THERAPIST", "C4",
        ) for day in ("ΔΕ", "ΤΡ", "ΤΕ", "ΠΕ", "ΠΑ")
    )
    plan = TrialImportPlan(
        sha256(review.read_bytes()).hexdigest(), plan.intake_sha256,
        1, 5, entries, plan.programme_sha256,
    )
    joined = project_trial_demo_sources(
        plan, programme=programme, review=review, intake=intake,
        clinical_approvals=_approval(),
    )
    assert joined.weekday_slot_count == 5
    assert [r[2] for r in joined.weekly_slots.rows] == [
        "ΔΕ", "ΤΡ", "ΤΕ", "ΠΕ", "ΠΑ",
    ]


def test_source_parity_rejects_correct_cell_reference_with_changed_time(tmp_path):
    from rehab_excel.trial_demo_source_reconciliation import (
        TrialDemoSourceReconciliationError, reconcile_demo_source_weekdays,
    )
    plan, programme, review, intake = _fixture(tmp_path)
    wb = load_workbook(programme)
    wb["ΠΡΟΓΡΑΜΜΑ"]["C4"] = "10:30"
    wb.save(programme)
    wb.close()
    updated = TrialImportPlan(
        plan.review_sha256, plan.intake_sha256, 1, 1, plan.entries,
        sha256(programme.read_bytes()).hexdigest(),
    )
    with pytest.raises(TrialDemoSourceReconciliationError, match="projection differs"):
        reconcile_demo_source_weekdays(
            updated, programme=programme, review=review, intake=intake,
        )


def test_source_parity_blocks_unmapped_afternoon_without_partial_schedule(tmp_path):
    from rehab_excel.trial_demo_source_reconciliation import (
        TrialDemoSourceReconciliationError, reconcile_demo_source_weekdays,
    )
    plan, programme, review, intake = _fixture(tmp_path)
    wb = load_workbook(programme)
    wb["ΠΡΟΓΡΑΜΜΑ"]["U4"] = "SYNTHETIC PM SERVICE"
    wb.save(programme)
    wb.close()
    updated = TrialImportPlan(
        plan.review_sha256, plan.intake_sha256, 1, 1, plan.entries,
        sha256(programme.read_bytes()).hexdigest(),
    )
    with pytest.raises(TrialDemoSourceReconciliationError, match="independent review"):
        reconcile_demo_source_weekdays(
            updated, programme=programme, review=review, intake=intake,
        )
