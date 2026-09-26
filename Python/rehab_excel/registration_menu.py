from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from enum import Enum
from typing import Mapping

from rehab_core.registration import (
    NewPatientRequest,
    NewStudentRequest,
    NewTherapistRequest,
)

from .registration_orchestrator import RegistrationKind, RegistrationRequest


class RegistrationMenuAction(str, Enum):
    NEW_PATIENT = "new_patient"
    NEW_THERAPIST = "new_therapist"
    NEW_STUDENT = "new_student"


class FormFieldType(str, Enum):
    TEXT = "text"
    INTEGER = "integer"
    DATE = "date"
    BOOLEAN = "boolean"


@dataclass(frozen=True)
class RegistrationFormField:
    key: str
    label: str
    field_type: FormFieldType
    required: bool = False
    default: object | None = None
    source: str | None = None
    help_text: str | None = None


@dataclass(frozen=True)
class RegistrationFormSpec:
    action: RegistrationMenuAction
    kind: RegistrationKind
    title: str
    fields: tuple[RegistrationFormField, ...]


REGISTRATION_FORMS: dict[RegistrationMenuAction, RegistrationFormSpec] = {
    RegistrationMenuAction.NEW_PATIENT: RegistrationFormSpec(
        action=RegistrationMenuAction.NEW_PATIENT,
        kind=RegistrationKind.PATIENT,
        title="Νέος ασθενής",
        fields=(
            RegistrationFormField(
                "patient_id",
                "Patient ID",
                FormFieldType.TEXT,
                required=True,
            ),
            RegistrationFormField(
                "display_name",
                "Ονοματεπώνυμο",
                FormFieldType.TEXT,
                required=True,
            ),
            RegistrationFormField(
                "room",
                "Θάλαμος",
                FormFieldType.TEXT,
                source="rooms",
            ),
            RegistrationFormField(
                "infectious",
                "Λοιμώδης ασθενής",
                FormFieldType.BOOLEAN,
                default=False,
            ),
            RegistrationFormField(
                "status",
                "Κατάσταση",
                FormFieldType.TEXT,
                source="patient_statuses",
            ),
        ),
    ),
    RegistrationMenuAction.NEW_THERAPIST: RegistrationFormSpec(
        action=RegistrationMenuAction.NEW_THERAPIST,
        kind=RegistrationKind.THERAPIST,
        title="Νέος θεραπευτής",
        fields=(
            RegistrationFormField(
                "display_name",
                "Ονοματεπώνυμο",
                FormFieldType.TEXT,
                required=True,
            ),
        ),
    ),
    RegistrationMenuAction.NEW_STUDENT: RegistrationFormSpec(
        action=RegistrationMenuAction.NEW_STUDENT,
        kind=RegistrationKind.STUDENT,
        title="Νέος φοιτητής",
        fields=(
            RegistrationFormField(
                "student_id",
                "Student ID",
                FormFieldType.TEXT,
                required=True,
            ),
            RegistrationFormField(
                "display_name",
                "Ονοματεπώνυμο",
                FormFieldType.TEXT,
                required=True,
            ),
            RegistrationFormField(
                "student_number",
                "Αριθμός φοιτητή",
                FormFieldType.INTEGER,
                required=True,
            ),
            RegistrationFormField(
                "placement_start",
                "Έναρξη πρακτικής",
                FormFieldType.DATE,
                required=True,
            ),
            RegistrationFormField(
                "placement_end",
                "Λήξη πρακτικής",
                FormFieldType.DATE,
                required=True,
            ),
            RegistrationFormField(
                "supervisor_therapist_id",
                "Επόπτης θεραπευτής",
                FormFieldType.TEXT,
                source="therapists",
            ),
            RegistrationFormField(
                "replacement_capable",
                "Δυνατότητα αναπληρώσεων",
                FormFieldType.BOOLEAN,
                default=True,
            ),
            RegistrationFormField(
                "robotic_capable",
                "Ρομποτική αποκατάσταση",
                FormFieldType.BOOLEAN,
                default=False,
            ),
        ),
    ),
}


def registration_menu_actions() -> tuple[RegistrationFormSpec, ...]:
    """Return the stable registration choices for a central menu."""

    return tuple(REGISTRATION_FORMS[action] for action in RegistrationMenuAction)


def registration_form_spec(
    action: RegistrationMenuAction | str,
) -> RegistrationFormSpec:
    try:
        resolved = (
            action
            if isinstance(action, RegistrationMenuAction)
            else RegistrationMenuAction(str(action).strip())
        )
    except ValueError as exc:
        raise ValueError(f"Unknown registration menu action: {action!r}") from exc
    return REGISTRATION_FORMS[resolved]


def _text(values: Mapping[str, object], key: str, *, required: bool = False) -> str | None:
    value = values.get(key)
    if value is None:
        if required:
            raise ValueError(f"{key} is required")
        return None
    text = str(value).strip()
    if required and not text:
        raise ValueError(f"{key} is required")
    return text or None


def _integer(values: Mapping[str, object], key: str) -> int:
    value = values.get(key)
    if value is None or str(value).strip() == "":
        raise ValueError(f"{key} is required")
    try:
        return int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{key} must be an integer") from exc


def _date(values: Mapping[str, object], key: str) -> date:
    value = values.get(key)
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value is None or not str(value).strip():
        raise ValueError(f"{key} is required")

    text = str(value).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"{key} must use DD/MM/YYYY or YYYY-MM-DD")


def _boolean(values: Mapping[str, object], key: str, default: bool) -> bool:
    value = values.get(key, default)
    if value is None or value == "":
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, int) and value in (0, 1):
        return bool(value)

    normalized = str(value).strip().casefold()
    if normalized in {"1", "true", "yes", "y", "ναι"}:
        return True
    if normalized in {"0", "false", "no", "n", "όχι", "οχι"}:
        return False
    raise ValueError(f"{key} must be a boolean value")


def build_registration_request(
    action: RegistrationMenuAction | str,
    values: Mapping[str, object],
) -> RegistrationRequest:
    """Translate menu/form values into the existing domain request objects.

    This layer performs only UI-shape parsing (text, dates, booleans, integers).
    Authoritative duplicate/configuration/business validation remains in the
    existing registration workflows.
    """

    resolved = registration_form_spec(action).action

    if resolved is RegistrationMenuAction.NEW_PATIENT:
        return NewPatientRequest(
            patient_id=_text(values, "patient_id", required=True) or "",
            display_name=_text(values, "display_name", required=True) or "",
            room=_text(values, "room"),
            infectious=_boolean(values, "infectious", False),
            status=_text(values, "status"),
        )

    if resolved is RegistrationMenuAction.NEW_THERAPIST:
        return NewTherapistRequest(
            display_name=_text(values, "display_name", required=True) or "",
            robotic_capable=False,
        )

    return NewStudentRequest(
        student_id=_text(values, "student_id", required=True) or "",
        display_name=_text(values, "display_name", required=True) or "",
        student_number=_integer(values, "student_number"),
        placement_start=_date(values, "placement_start"),
        placement_end=_date(values, "placement_end"),
        supervisor_therapist_id=_text(values, "supervisor_therapist_id"),
        replacement_capable=_boolean(values, "replacement_capable", True),
        robotic_capable=_boolean(values, "robotic_capable", False),
    )
