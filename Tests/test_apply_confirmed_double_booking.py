from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
TOOL_PATH = REPO_ROOT / "Python" / "tools" / "apply_confirmed_double_booking.py"


def test_apply_bridge_requires_explicit_confirmation_and_review_membership():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert 'if not args.confirm:' in source
    assert 'explicit --confirm is required for double booking' in source
    assert '_load_review_choice(' in source
    assert 'requires_user_confirmation' in source
    assert 'Selected therapist/time is not present in review options' in source


def test_apply_bridge_revalidates_normal_unresolved_before_exception():
    source = TOOL_PATH.read_text(encoding="utf-8")

    normal_plan = source.index("normal_result = build_global_replacement_assignments(")
    unresolved_guard = source.index(
        "if args.target_session_id not in normal_result.unresolved_request_ids:"
    )
    proposal = source.index("proposal = propose_double_booking_exception(")
    confirmation = source.index("assignment = confirm_double_booking_exception(")

    assert normal_plan < unresolved_guard < proposal < confirmation


def test_apply_bridge_keeps_fresh_preview_mode_for_native_regression():
    source = TOOL_PATH.read_text(encoding="utf-8")

    copy_step = source.index("shutil.copy2(input_book, output)")
    persist_step = source.index("append_daily_replacement_facts(output")
    output_contract = source.index("output_contract = assert_application_contract(output)")

    assert copy_step < persist_step < output_contract
    assert "read_daily_replacement_facts(" in source
    assert "Confirmed double-booking fact did not read back exactly once" in source
    assert "file_sha256(input_book) == source_hash" in source


def test_apply_bridge_uses_registry_aware_daily_input_validation():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert "therapist_names=settings.therapist_names" in source
    assert 'reason="CONFIRMED DOUBLE BOOKING EXCEPTION"' not in source
    assert "confirm_double_booking_exception(" in source


def test_apply_bridge_validate_only_returns_complete_verified_assignment_without_output():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert 'parser.add_argument("--validate-only", action="store_true")' in source
    assert 'if not args.validate_only and args.output is None:' in source
    assert '"validated_only": args.validate_only' in source
    assert '"assignment": _assignment_payload(assignment)' in source
    assert '"replacement_id": assignment.replacement_id' in source
    assert '"replacement_therapist_id": assignment.replacement_therapist_id' in source
    assert '"status": "ACTIVE"' in source
    assert 'if not args.validate_only:' in source
    assert "CONFIRMED DOUBLE BOOKING VALIDATION OK" in source
