from datetime import date, time

from rehab_excel.patient_centric_preview import _temporary_date_label


def compact_source_label(provider_name: str, slot: time) -> str:
    # Mirrors the intentionally compact source-cell convention.
    return f"→ {provider_name} {slot.strftime('%H:%M')}"


def test_source_replacement_note_has_no_repeated_date():
    label = compact_source_label("Καλυβιώτης", time(9, 15))
    assert label == "→ Καλυβιώτης 09:15"
    assert "18/09" not in label


def test_temporary_date_label_is_still_available_for_destination_assignment():
    assert _temporary_date_label(date(2026, 9, 18)) == "Πα 18/09"
