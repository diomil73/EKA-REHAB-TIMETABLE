from rehab_excel import (
    AUTHORITATIVE_SHEETS,
    SheetAuthority,
    get_source_rule,
    imported_sheets,
    is_authoritative,
)


def test_only_confirmed_input_sheets_are_authoritative():
    assert AUTHORITATIVE_SHEETS == {"PATIENTS", "PATIENT_PLANNER", "SETTINGS"}
    assert imported_sheets() == AUTHORITATIVE_SHEETS


def test_sessions_is_explicitly_audit_only():
    rule = get_source_rule("SESSIONS")
    assert rule is not None
    assert rule.authority is SheetAuthority.AUDIT_ONLY
    assert rule.python_import is False
    assert is_authoritative("SESSIONS") is False


def test_master_and_daily_are_views_not_sources():
    assert get_source_rule("MASTER_SCHEDULE").authority is SheetAuthority.DERIVED_VIEW
    assert get_source_rule("THERAPIST_DAILY").authority is SheetAuthority.DERIVED_VIEW
    assert is_authoritative("MASTER_SCHEDULE") is False


def test_existing_operational_sheets_are_not_silently_promoted_to_sources():
    assert get_source_rule("THERAPIST_ATTENDANCE").authority is SheetAuthority.LEGACY_OPERATIONAL
    assert get_source_rule("REPLACEMENT_LOG").authority is SheetAuthority.LEGACY_OPERATIONAL
