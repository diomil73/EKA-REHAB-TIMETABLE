import inspect

from tools import prepare_global_replacements_in_place as module


def test_global_replacement_preparation_is_read_only_and_uses_authoritative_solver():
    source = inspect.getsource(module.prepare_global_replacements_payload)

    assert "assert_application_contract(source)" in source
    assert "build_therapist_absence_replacement_queue" in source
    assert "build_global_replacement_assignments" in source
    assert "build_double_booking_review_items" in source
    assert "build_daily_replacement_fact_rows" in source
    assert "file_sha256(source)" in source

    assert "apply_write_plan_to_copy" not in source
    assert "Workbooks.Open" not in source
    assert "shutil.copy" not in source


def test_global_replacement_payload_preserves_explicit_double_booking_review():
    source = inspect.getsource(module.prepare_global_replacements_payload)

    assert 'mode = "review_required"' in source
    assert "reviewable = tuple(item for item in review_items if item.has_options)" in source
    assert '"review_required": bool(reviewable)' in source
    assert "build_unresolved_suspensions" in source


def test_resolved_global_replacement_plan_updates_therapist_and_master_views():
    source = inspect.getsource(module.prepare_global_replacements_payload)

    assert "build_patient_centric_preview_plan" in source
    assert "compose_outpatient_daily_plan" in source
    assert "build_master_daily_overlay_plan" in source
    assert "compose_master_daily_overlay" in source
    assert "validate_write_plan(plan)" in source
    assert 'lines.append(_patch_line(patch))' in source


def test_global_payload_contains_replacement_and_suspension_records():
    source = inspect.getsource(module.prepare_global_replacements_payload)

    assert '"REPLACEMENT"' in source
    assert '"SUSPENSION"' in source
    assert "row.provider_kind.value" in source
    assert "SUSPENSION_STATUS" in source
    assert "SUSPENSION_REASON" in source
