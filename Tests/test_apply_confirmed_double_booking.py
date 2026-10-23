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


def test_apply_bridge_writes_only_fresh_preview_fact_via_excel_com_path():
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


def test_apply_bridge_can_return_authoritative_commit_metadata():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert 'parser.add_argument("--response", type=Path)' in source
    assert "def _write_response(" in source
    assert '"source_sha256_before": source_hash' in source
    assert '"output_path": str(output)' in source
    assert '"ok": True' in source
    assert '"ok": False' in source
    assert "output.unlink(missing_ok=True)" in source
