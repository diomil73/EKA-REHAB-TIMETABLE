from __future__ import annotations

from typing import Iterable

from .models import Patient


def next_sequential_patient_id(existing_patients: Iterable[Patient]) -> str:
    """Return the next numeric internal PatientID without reusing existing IDs.

    Legacy/non-numeric IDs are preserved and ignored for sequencing. Numeric
    IDs are compared by integer value, so the next ID is always greater than
    every existing numeric PatientID.
    """

    numeric_ids: list[int] = []
    for patient in existing_patients:
        value = str(patient.patient_id).strip()
        if value.isdigit():
            numeric_ids.append(int(value))
    return str(max(numeric_ids, default=0) + 1)
