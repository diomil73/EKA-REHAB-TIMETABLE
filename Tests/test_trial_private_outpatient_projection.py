"""Fictional-only source projection regression. No real clinical records."""
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from rehab_excel.trial_private_outpatient_projection import (
    TrialPrivateOutpatientProjectionError, project_private_outpatient_source,
)


def _source(tmp_path: Path) -> Path:
    workbook = Workbook()
    patients = workbook.active
    patients.title = "PATIENTS"
    patients.append((
        "PatientID", "Θάλαμος", "Ασθενής", "Λοιμώδης",
        "Κατάσταση", "PatientType", "HospitalMRN",
        "ResponsibleDoctor", "SourceContext",
    ))
    patients.append((
        "OUT-1", None, "FICTIONAL PATIENT", "Ο",
        "Ενεργός", "Εξωτερικός", None, None, "APPROVED SYNTHETIC",
    ))
    schedule = workbook.create_sheet("OUTPATIENT_SCHEDULE")
    schedule.append((
        "PatientID", "Ασθενής", "Θεραπεία",
        "Ώρα", "Ημέρες", "Θεραπευτής",
    ))
    schedule.append(("OUT-1", "FICTIONAL PATIENT", "ΦΘ", "09:15", "Δε-Τε-Πα", "DEMO T"))
    schedule.append(("OUT-1", "FICTIONAL PATIENT", "Εργο", "10:15", "Τρ-Πε", None))
    path = tmp_path / "private_synthetic_outpatients.xlsx"
    workbook.save(path)
    workbook.close()
    return path


def _modify(path, updater):
    workbook = load_workbook(path)
    updater(workbook)
    workbook.save(path)
    workbook.close()


def test_projection_preserves_existing_application_fields_and_exact_weekdays(tmp_path):
    source = _source(tmp_path)
    original = source.read_bytes()
    result = project_private_outpatient_source(source)
    assert result.patient_count == 1
    assert result.appointment_count == 2
    assert result.weekday_slot_count == 5
    assert result.patient_rows[0][-1] == "APPROVED SYNTHETIC"
    assert result.appointment_rows[0][:3] == ("OUT-1", "FICTIONAL PATIENT", "ΦΘ")
    assert {r[3] for r in result.weekday_view_rows} == {"ΔΕ", "ΤΡ", "ΤΕ", "ΠΕ", "ΠΑ"}
    assert all(r[6].startswith("OUTPATIENT_SCHEDULE:") for r in result.weekday_view_rows)
    assert source.read_bytes() == original
    assert "FICTIONAL PATIENT" not in repr(result)


def test_populated_unknown_metadata_fails_closed(tmp_path):
    source = _source(tmp_path)
    _modify(source, lambda wb: wb["PATIENTS"].cell(2, 10, "UNAPPROVED DETAIL"))
    with pytest.raises(TrialPrivateOutpatientProjectionError):
        project_private_outpatient_source(source)
    source = _source(tmp_path)
    _modify(source, lambda wb: wb["OUTPATIENT_SCHEDULE"].cell(2, 7, "PRIVATE NOTE"))
    with pytest.raises(TrialPrivateOutpatientProjectionError):
        project_private_outpatient_source(source)


def test_readonly_source_rejected_in_git_checkout(tmp_path):
    source = _source(tmp_path)
    (tmp_path / ".git").mkdir()
    with pytest.raises(TrialPrivateOutpatientProjectionError):
        project_private_outpatient_source(source)


def test_patient_type_mismatch_is_not_imported(tmp_path):
    source = _source(tmp_path)
    _modify(source, lambda wb: wb["PATIENTS"].cell(2, 6, "Εσωτερικός"))
    with pytest.raises(TrialPrivateOutpatientProjectionError):
        project_private_outpatient_source(source)


def _inpatient_projection(*, day="ΔΕ", clock=None, provider="DEMO T", approved_shared=False):
    from datetime import time
    from rehab_excel.trial_demo_source_projection import TrialDemoSourceProjection
    from rehab_excel.trial_patient_registry_projection import (
        PATIENTS_HEADERS, TrialPatientRegistryProjection,
    )
    from rehab_excel.trial_weekly_projection import TrialWeeklyProjection
    from rehab_excel.weekly_slots_source import HEADERS

    start = clock or time(9, 15)
    patients = [("IN-1", None, "FICTIONAL INPATIENT A", "Ο", "Ενεργός", "Εσωτερικός", None, None, None)]
    weekly = [("IN-1", "ΦΘ", day, start, provider, "C2")]
    if approved_shared:
        patients.append(("IN-2", None, "FICTIONAL INPATIENT B", "Ο", "Ενεργός", "Εσωτερικός", None, None, None))
        weekly.append(("IN-2", "ΦΘ", day, start, provider, "C3"))
    return TrialDemoSourceProjection(
        TrialPatientRegistryProjection(PATIENTS_HEADERS, tuple(patients), len(patients)),
        TrialWeeklyProjection(HEADERS, tuple(weekly), len(weekly)),
        len(patients), len(weekly),
    )


@pytest.mark.parametrize("provider", ["DEMO T", " demo t "])
def test_mixed_patient_types_cannot_share_provider_slot_without_approval(tmp_path, provider):
    from rehab_excel.trial_demo_source_projection import merge_trial_demo_outpatients

    source = _source(tmp_path)
    before = source.read_bytes()
    outpatient = project_private_outpatient_source(source)
    with pytest.raises(ValueError, match="explicit shared-slot approval") as err:
        merge_trial_demo_outpatients(_inpatient_projection(provider=provider), outpatient)
    assert "FICTIONAL" not in str(err.value)
    assert source.read_bytes() == before


@pytest.mark.parametrize("field", ["day", "clock", "provider"])
def test_mixed_sources_with_distinct_provider_slots_remain_complete(tmp_path, field):
    from datetime import time
    from rehab_excel.trial_demo_source_projection import merge_trial_demo_outpatients

    options = {"day": "ΤΡ", "clock": time(9, 45), "provider": "OTHER DEMO T"}
    inpatient = _inpatient_projection(**{field: options[field]})
    outpatient = project_private_outpatient_source(_source(tmp_path))
    combined = merge_trial_demo_outpatients(inpatient, outpatient)
    assert combined.patients == 2
    assert combined.weekday_slot_count == 6
    assert combined.weekly_slots == inpatient.weekly_slots
    assert combined.outpatient is outpatient


def test_outpatient_pair_cannot_implicitly_share_one_provider_slot(tmp_path):
    from rehab_excel.trial_demo_source_projection import merge_trial_demo_outpatients

    source = _source(tmp_path)
    def add_peer(wb):
        wb["PATIENTS"].append(("OUT-2", None, "FICTIONAL PEER", "Ο", "Ενεργός", "Εξωτερικός"))
        wb["OUTPATIENT_SCHEDULE"].append(("OUT-2", "FICTIONAL PEER", "ΦΘ", "09:15", "Τε", "DEMO T"))
    _modify(source, add_peer)
    outpatient = project_private_outpatient_source(source)
    with pytest.raises(ValueError, match="explicit shared-slot approval"):
        merge_trial_demo_outpatients(_inpatient_projection(day="ΤΡ"), outpatient)


def test_existing_approved_inpatient_group_is_not_rejected_or_expanded(tmp_path):
    from rehab_excel.trial_demo_source_projection import merge_trial_demo_outpatients

    outpatient = project_private_outpatient_source(_source(tmp_path))
    existing = _inpatient_projection(day="ΤΡ", approved_shared=True)
    combined = merge_trial_demo_outpatients(existing, outpatient)
    assert combined.patients == 3
    assert combined.weekday_slot_count == 7
    assert combined.weekly_slots.rows == existing.weekly_slots.rows
    with pytest.raises(ValueError, match="explicit shared-slot approval"):
        merge_trial_demo_outpatients(
            _inpatient_projection(approved_shared=True), outpatient,
        )


def test_mixed_source_merge_uses_existing_greek_provider_label_normalization(tmp_path):
    from rehab_excel.trial_demo_source_projection import merge_trial_demo_outpatients

    source = _source(tmp_path)
    _modify(source, lambda wb: setattr(wb["OUTPATIENT_SCHEDULE"]["F2"], "value", "Δοκιμή. Α"))
    outpatient = project_private_outpatient_source(source)
    with pytest.raises(ValueError, match="explicit shared-slot approval"):
        merge_trial_demo_outpatients(_inpatient_projection(provider="δοκιμηα"), outpatient)


def test_capacity_audit_counts_distinct_clock_slots_not_patients(tmp_path):
    from rehab_excel.trial_demo_source_projection import merge_trial_demo_outpatients
    from rehab_excel.trial_provider_capacity import audit_trial_provider_capacity

    inpatient = _inpatient_projection(day="ΤΡ", approved_shared=True)
    outpatient = project_private_outpatient_source(_source(tmp_path))
    combined = merge_trial_demo_outpatients(inpatient, outpatient)
    report = audit_trial_provider_capacity(combined, approved_limits={
        ("demo.t", day): 1 for day in ("ΔΕ", "ΤΡ", "ΤΕ", "ΠΑ")
    })
    assert report.named_provider_capacity_clear
    assert report.named_provider_placements == 5
    assert report.occupied_distinct_slots == 4
    assert report.unassigned_placements == 2
    assert "DEMO" not in repr(report)


def test_capacity_audit_uses_actual_day_limit_and_no_default(tmp_path):
    from datetime import time
    from rehab_excel.trial_demo_source_projection import merge_trial_demo_outpatients
    from rehab_excel.trial_provider_capacity import audit_trial_provider_capacity

    outpatient = project_private_outpatient_source(_source(tmp_path))
    combined = merge_trial_demo_outpatients(_inpatient_projection(clock=time(9, 45)), outpatient)
    report = audit_trial_provider_capacity(combined, approved_limits={
        ("DEMO T", "ΔΕ"): 1, ("DEMO T", "ΤΕ"): 0,
    })
    assert not report.named_provider_capacity_clear
    assert report.exceeded_limit_days == 2
    assert report.missing_limit_days == 1
    assert report.occupied_distinct_slots == 4


@pytest.mark.parametrize("limits", [
    {("DEMO T", "ΔΕ"): True},
    {("DEMO T", "ΔΕ"): -1},
    {("DEMO T", "ΣΑ"): 6},
    {("DEMO T", "ΔΕ"): 1, ("demo.t", "ΔΕ"): 2},
])
def test_capacity_audit_rejects_ambiguous_or_invalid_limits(limits):
    from rehab_excel.trial_provider_capacity import audit_trial_provider_capacity, TrialProviderCapacityError

    with pytest.raises(TrialProviderCapacityError):
        audit_trial_provider_capacity(_inpatient_projection(), approved_limits=limits)


def _provider_snapshot(tmp_path, outpatient, *, shared=None, monday=None):
    import json
    from types import SimpleNamespace
    from rehab_excel.trial_provider_schedule import SCHEMA
    plan = SimpleNamespace(programme_sha256="a"*64, review_sha256="b"*64, intake_sha256="c"*64)
    path = tmp_path / "PRIVATE_PROVIDER_SCHEDULE.json"
    slots = [{"provider": "DEMO T", "weekday": day, "times": ["09:15"], "decision": "APPROVED"}
             for day in ("ΔΕ", "ΤΡ", "ΤΕ", "ΠΕ", "ΠΑ")]
    if monday is not None:
        slots[0]["times"] = monday
    data = {"schema": SCHEMA,
            "sources": {"programme": plan.programme_sha256, "review": plan.review_sha256,
                        "shell": plan.intake_sha256, "outpatients": outpatient.source_sha256},
            "availability": slots, "shared_slots": shared or []}
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return plan, path


def test_provider_snapshot_preserves_complementary_days_as_one_clock_per_day(tmp_path):
    from rehab_excel.trial_provider_schedule import read_private_provider_schedule, merge_with_approved_provider_schedule

    outpatient = project_private_outpatient_source(_source(tmp_path))
    inpatient = _inpatient_projection(day="ΤΡ")
    plan, path = _provider_snapshot(tmp_path, outpatient)
    result = merge_with_approved_provider_schedule(
        inpatient, outpatient, plan=plan, schedule=read_private_provider_schedule(path)
    )
    assert result.patients == 2
    assert result.weekday_slot_count == 6


def test_exact_mixed_membership_can_be_explicitly_approved(tmp_path):
    from rehab_excel.trial_provider_schedule import read_private_provider_schedule, merge_with_approved_provider_schedule

    outpatient = project_private_outpatient_source(_source(tmp_path))
    shared = [{"provider": "DEMO T", "weekday": "ΔΕ", "time": "09:15",
               "patient_ids": ["IN-1", "OUT-1"], "decision": "APPROVED_SHARED"}]
    plan, path = _provider_snapshot(tmp_path, outpatient, shared=shared)
    before = path.read_bytes()
    result = merge_with_approved_provider_schedule(
        _inpatient_projection(), outpatient, plan=plan, schedule=read_private_provider_schedule(path),
    )
    assert result.patients == 2 and result.weekday_slot_count == 6
    assert path.read_bytes() == before


def test_partial_mixed_membership_cannot_extend_approved_group(tmp_path):
    from rehab_excel.trial_provider_schedule import read_private_provider_schedule, merge_with_approved_provider_schedule

    outpatient = project_private_outpatient_source(_source(tmp_path))
    shared = [{"provider": "DEMO T", "weekday": "ΔΕ", "time": "09:15",
               "patient_ids": ["IN-1", "OUT-1"], "decision": "APPROVED_SHARED"}]
    plan, path = _provider_snapshot(tmp_path, outpatient, shared=shared)
    with pytest.raises(ValueError, match="explicit shared-slot approval"):
        merge_with_approved_provider_schedule(
            _inpatient_projection(approved_shared=True), outpatient,
            plan=plan, schedule=read_private_provider_schedule(path),
        )


def test_stale_provider_snapshot_cannot_authorize_changed_source(tmp_path):
    from rehab_excel.trial_provider_schedule import (
        read_private_provider_schedule, merge_with_approved_provider_schedule, TrialProviderScheduleError,
    )

    outpatient = project_private_outpatient_source(_source(tmp_path))
    plan, path = _provider_snapshot(tmp_path, outpatient)
    schedule = read_private_provider_schedule(path)
    plan.review_sha256 = "d"*64
    with pytest.raises(TrialProviderScheduleError, match="different sources"):
        merge_with_approved_provider_schedule(_inpatient_projection(day="ΤΡ"), outpatient, plan=plan, schedule=schedule)


def test_available_clock_not_just_number_of_available_slots(tmp_path):
    from rehab_excel.trial_provider_schedule import (
        read_private_provider_schedule, merge_with_approved_provider_schedule, TrialProviderScheduleError,
    )
    outpatient = project_private_outpatient_source(_source(tmp_path))
    plan, path = _provider_snapshot(tmp_path, outpatient, monday=["10:15"])
    with pytest.raises(TrialProviderScheduleError, match="not explicitly available"):
        merge_with_approved_provider_schedule(
            _inpatient_projection(day="ΤΡ"), outpatient, plan=plan, schedule=read_private_provider_schedule(path)
        )


def test_stale_shared_decision_not_reused_for_complementary_days(tmp_path):
    from rehab_excel.trial_provider_schedule import read_private_provider_schedule, merge_with_approved_provider_schedule

    outpatient = project_private_outpatient_source(_source(tmp_path))
    shared = [{"provider": "DEMO T", "weekday": "ΔΕ", "time": "09:15",
               "patient_ids": ["IN-1", "OUT-1"], "decision": "APPROVED_SHARED"}]
    plan, path = _provider_snapshot(tmp_path, outpatient, shared=shared)
    with pytest.raises(ValueError, match="Stale or mismatched"):
        merge_with_approved_provider_schedule(
            _inpatient_projection(day="ΤΡ"), outpatient, plan=plan, schedule=read_private_provider_schedule(path)
        )


@pytest.mark.parametrize("decision", ["PENDING", "", True])
def test_availability_requires_literal_explicit_approval(tmp_path, decision):
    import json
    from rehab_excel.trial_provider_schedule import read_private_provider_schedule, TrialProviderScheduleError
    outpatient = project_private_outpatient_source(_source(tmp_path))
    _, path = _provider_snapshot(tmp_path, outpatient)
    data = json.loads(path.read_text(encoding="utf-8"))
    data["availability"][0]["decision"] = decision
    path.write_text(json.dumps(data), encoding="utf-8")
    with pytest.raises(TrialProviderScheduleError, match="Explicit availability"):
        read_private_provider_schedule(path)
