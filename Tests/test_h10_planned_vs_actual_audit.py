from datetime import date
from types import SimpleNamespace

import pytest

from tools.audit_therapist_planned_vs_actual import positive_deltas


def daily(day, therapist, planned, delivered, **extra):
    return SimpleNamespace(
        service_date=date.fromisoformat(day),
        therapist_id=therapist,
        planned_sessions=planned,
        final_active_sessions=delivered,
        replacements_received=extra.get("replacements", 0),
        approved_shared_or_over_capacity_events=extra.get("approved_extra", 0),
        released_by_cancellation_or_no_show=extra.get("released", 0),
        treatment_suspensions_affecting_load=extra.get("suspensions", 0),
    )


def test_detects_only_actual_over_original_planned_without_making_causal_claims():
    rows = [
        daily("2026-10-30", "Argentos", 3, 5, replacements=1, approved_extra=2),
        daily("2026-10-29", "Argentos", 5, 3),
        daily("2026-10-29", "Someone", 3, 4),
    ]
    flagged = positive_deltas(rows)
    assert [(r["date"], r["therapist"], r["delta"]) for r in flagged] == [
        ("2026-10-29", "Someone", 1),
        ("2026-10-30", "Argentos", 2),
    ]
    assert flagged[1]["initial_planned"] == 3
    assert flagged[1]["final_delivered"] == 5
    assert flagged[1]["replacements"] == 1
    assert flagged[1]["approved_extra"] == 2
    assert positive_deltas(rows, "argentOS") == [flagged[1]]


def test_duplicate_and_negative_counts_fail_closed():
    d = daily("2026-10-30", "T1", 3, 5)
    with pytest.raises(ValueError, match="Duplicate"):
        positive_deltas([d, d])
    with pytest.raises(ValueError, match="Negative"):
        positive_deltas([daily("2026-10-30", "T1", -1, 0)])
