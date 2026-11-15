"""Synthetic dry-run tests; no Excel or real clinical data."""
from datetime import time

import pytest

from rehab_excel.afternoon_slots_source import AfternoonSlot
from rehab_excel.afternoon_patient_edit_plan import (
    AfternoonEditError,
    prepare_afternoon_patient_edit,
)


PATIENTS = ("SYNTH-P1", "SYNTH-P2", "SYNTH-P3")


def row(pid, day="ΔΕ", hour=16, therapist="SYNTH-PT"):
    return AfternoonSlot(pid, "ΦΘ", day, time(hour), therapist)


def plan(current, patient, replacements, **kw):
    return prepare_afternoon_patient_edit(
        current, patient_id=patient, replacement_rows=replacements,
        known_patient_ids=PATIENTS, **kw
    )


def test_add_only_selected_patient_and_preserve_existing_rows():
    previous = (row("SYNTH-P2"),)
    update = plan(previous, "SYNTH-P1", (row("SYNTH-P1"),))
    assert update.after == (row("SYNTH-P2"), row("SYNTH-P1"))
    assert update.before == previous
    assert (update.added_count, update.removed_count) == (1, 0)


def test_modify_only_one_patient_maintains_other_rows_order():
    existing = (
        row("SYNTH-P2"), row("SYNTH-P1"), row("SYNTH-P3"),
        row("SYNTH-P1", "ΤΡ", 17), row("SYNTH-P2", "ΤΕ", 18),
    )
    new = (row("SYNTH-P1", "ΠΕ", 18, "SYNTH-OTHER"),)
    edit = plan(existing, "SYNTH-P1", new)
    assert edit.after == (
        row("SYNTH-P2"), *new, row("SYNTH-P3"), row("SYNTH-P2", "ΤΕ", 18),
    )
    assert edit.removed_count == 2
    assert edit.added_count == 1
    assert edit.unchanged_count == 0
    assert existing[1] == row("SYNTH-P1")


def test_dry_run_is_idempotent():
    existing = (row("SYNTH-P1"), row("SYNTH-P2"))
    first = plan(existing, "SYNTH-P1", (row("SYNTH-P1"),))
    second = plan(first.after, "SYNTH-P1", (row("SYNTH-P1"),))
    assert first.after == second.after == existing
    assert not first.changes_required
    assert second.added_count == second.removed_count == 0
    assert second.unchanged_count == 1


def test_empty_edit_requires_explicit_permission_to_remove():
    existing = (row("SYNTH-P1"), row("SYNTH-P2"))
    with pytest.raises( AfternoonEditError):
        plan(existing, "SYNTH-P1", ())
    cleared = plan(existing, "SYNTH-P1", (), allow_clear=True)
    assert cleared.after == (row("SYNTH-P2"),)
    assert cleared.removed_count == 1


def test_blank_new_patient_stays_blank():
    existing = (row("SYNTH-P2"),)
    updated = plan(existing, "SYNTH-P1", ())
    assert updated.after == existing
    assert not updated.changes_required


def test_cannot_modify_other_patient_in_replacement():
    with pytest.raises(AfternoonEditError):
        plan((row("SYNTH-P1"),), "SYNTH-P1", (row("SYNTH-P2"),))


def test_reject_unconfirmed_patient_id():
    with pytest.raises(AfternoonEditError):
        plan((), "SYNTH-UNKNOWN", (row("SYNTH-UNKNOWN"),))


def test_reject_duplicate_confirmed_patient_ids():
    with pytest.raises(AfternoonEditError):
        prepare_afternoon_patient_edit(
            (), patient_id="SYNTH-P1", replacement_rows=(),
            known_patient_ids=("SYNTH-P1", "SYNTH-P1"),
        )


def test_existing_invalid_rows_are_not_silently_repaired():
    with pytest.raises(AfternoonEditError):
        plan((row("SYNTH-P1"), row("SYNTH-P1")), "SYNTH-P1", (row("SYNTH-P1"),))


def test_reject_duplicate_replacement_rows():
    with pytest.raises(AfternoonEditError):
        plan((), "SYNTH-P1", (row("SYNTH-P1"), row("SYNTH-P1")))


def test_different_patient_shared_therapist_requires_separate_later_review():
    current = (row("SYNTH-P2"),)
    proposal = plan(current, "SYNTH-P1", (row("SYNTH-P1"),))
    assert len(proposal.after) == 2  # Not authorization for shared sessions


def test_preserves_afternoon_only_namespace_by_type():
    proposal = plan((), "SYNTH-P1", (row("SYNTH-P1"),))
    assert all(item.treatment == "ΦΘ" for item in proposal.after)
    assert all(item.patient_id == "SYNTH-P1" for item in proposal.after)
