from datetime import date, time
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "Python"))

from rehab_core.models import Session, Therapist
from rehab_core.provider_policy import (
    LeadershipAssignment,
    LeadershipRole,
    OverrideKind,
    ProviderOverride,
    ProviderPolicyBook,
    ProviderProfile,
)
from rehab_core.replacement_policy import find_policy_replacement_options


TODAY = date(2026, 9, 24)
SLOT_0915 = time(9, 15)
SLOT_1045 = time(10, 45)


def target() -> Session:
    return Session("TARGET", "P1", "ORIGINAL", TODAY, SLOT_0915)


def book_for(*providers: str) -> ProviderPolicyBook:
    book = ProviderPolicyBook()
    for provider in providers:
        book.profiles[provider] = ProviderProfile(provider, provider)
    return book


def test_excluded_provider_is_not_ranked():
    book = book_for("A", "B")
    book.overrides.append(
        ProviderOverride("A", OverrideKind.EXCLUDE_REPLACEMENTS, TODAY, TODAY)
    )
    options = find_policy_replacement_options(
        target(),
        therapists=[Therapist("A", "A"), Therapist("B", "B")],
        sessions=[],
        policy_book=book,
        timeslots=[SLOT_0915],
    )
    assert [option.provider_id for option in options] == ["B"]


def test_temporary_max_load_is_enforced_by_live_ranking():
    book = book_for("A")
    book.overrides.append(
        ProviderOverride("A", OverrideKind.MAX_LOAD, TODAY, TODAY, max_load=1)
    )
    sessions = [Session("A1", "P2", "A", TODAY, SLOT_1045)]
    options = find_policy_replacement_options(
        target(),
        therapists=[Therapist("A", "A")],
        sessions=sessions,
        policy_book=book,
        timeslots=[SLOT_0915, SLOT_1045],
    )
    assert options == []


def test_blocked_time_is_removed_but_other_free_time_remains():
    book = book_for("A")
    book.overrides.append(
        ProviderOverride(
            "A", OverrideKind.BLOCK_TIME, TODAY, TODAY, time_slot=SLOT_0915
        )
    )
    options = find_policy_replacement_options(
        target(),
        therapists=[Therapist("A", "A")],
        sessions=[],
        policy_book=book,
        timeslots=[SLOT_0915, SLOT_1045],
    )
    assert len(options) == 1
    assert options[0].recommended_time == SLOT_1045
    assert options[0].available_timeslots == (SLOT_1045,)
    assert options[0].exact_time_available is False


def test_manager_appears_only_at_explicitly_open_replacement_time():
    book = book_for("MANAGER")
    book.leadership.append(
        LeadershipAssignment("MANAGER", LeadershipRole.MANAGER, TODAY)
    )
    book.overrides.append(
        ProviderOverride(
            "MANAGER",
            OverrideKind.OPEN_REPLACEMENT_TIME,
            TODAY,
            TODAY,
            time_slot=SLOT_1045,
        )
    )
    options = find_policy_replacement_options(
        target(),
        therapists=[Therapist("MANAGER", "MANAGER")],
        sessions=[],
        policy_book=book,
        timeslots=[SLOT_0915, SLOT_1045],
    )
    assert len(options) == 1
    assert options[0].available_timeslots == (SLOT_1045,)
    assert options[0].recommended_time == SLOT_1045


def test_manager_with_no_open_times_is_hidden():
    book = book_for("MANAGER")
    book.leadership.append(
        LeadershipAssignment("MANAGER", LeadershipRole.MANAGER, TODAY)
    )
    options = find_policy_replacement_options(
        target(),
        therapists=[Therapist("MANAGER", "MANAGER")],
        sessions=[],
        policy_book=book,
        timeslots=[SLOT_0915, SLOT_1045],
    )
    assert options == []
