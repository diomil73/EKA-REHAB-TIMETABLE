from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Iterable
import unicodedata

from .models import Patient, PatientType, Student, Therapist


@dataclass(frozen=True)
class RegistrationIssue:
    code: str
    field: str
    message: str


@dataclass(frozen=True)
class RegistrationCheck:
    issues: tuple[RegistrationIssue, ...] = ()

    @property
    def allowed(self) -> bool:
        return not self.issues


@dataclass(frozen=True)
class NewPatientRequest:
    patient_id: str
    display_name: str
    room: str | None = None
    infectious: bool = False
    status: str | None = None
    patient_type: PatientType = PatientType.INPATIENT
    hospital_mrn: str | None = None


@dataclass(frozen=True)
class NewTherapistRequest:
    display_name: str
    robotic_capable: bool = False


@dataclass(frozen=True)
class NewStudentRequest:
    student_id: str
    display_name: str
    student_number: int
    placement_start: date
    placement_end: date
    supervisor_therapist_id: str | None = None
    replacement_capable: bool = True
    robotic_capable: bool = False


def _key(value: object) -> str:
    """Return a comparison key that is whitespace/case/accent insensitive.

    Greek registries commonly mix accented and unaccented uppercase/lowercase
    spellings (for example ``Πέτσιος`` and ``ΠΕΤΣΙΟΣ``). Those must resolve to
    the same logical provider when checking duplicates.
    """

    text = unicodedata.normalize("NFD", str(value).strip().casefold())
    return "".join(char for char in text if not unicodedata.combining(char))


def _nonblank(value: object) -> bool:
    return bool(str(value).strip())


def validate_new_patient(
    request: NewPatientRequest,
    *,
    existing_patients: Iterable[Patient] = (),
    allowed_statuses: Iterable[str] | None = None,
    allowed_rooms: Iterable[str] | None = None,
) -> RegistrationCheck:
    issues: list[RegistrationIssue] = []

    if not _nonblank(request.patient_id):
        issues.append(RegistrationIssue("patient_id_required", "patient_id", "PatientID is required"))
    elif _key(request.patient_id) in {_key(item.patient_id) for item in existing_patients}:
        issues.append(RegistrationIssue("duplicate_patient_id", "patient_id", "PatientID already exists"))

    if not _nonblank(request.display_name):
        issues.append(RegistrationIssue("patient_name_required", "display_name", "Patient name is required"))

    if request.patient_type == PatientType.OUTPATIENT:
        if request.infectious:
            issues.append(
                RegistrationIssue(
                    "outpatient_infectious_not_allowed",
                    "infectious",
                    "Outpatient patient cannot be infectious",
                )
            )
        if request.room is not None and str(request.room).strip():
            issues.append(
                RegistrationIssue(
                    "outpatient_room_not_allowed",
                    "room",
                    "Outpatient patient must not have an inpatient room",
                )
            )

    if request.status is not None and allowed_statuses is not None:
        allowed = {_key(value) for value in allowed_statuses}
        if _key(request.status) not in allowed:
            issues.append(RegistrationIssue("invalid_patient_status", "status", "Patient status is not in the configured list"))

    if request.room is not None and allowed_rooms is not None:
        allowed = {_key(value) for value in allowed_rooms}
        if _key(request.room) not in allowed:
            issues.append(RegistrationIssue("invalid_room", "room", "Room is not in the configured list"))

    return RegistrationCheck(tuple(issues))


def validate_new_therapist(
    request: NewTherapistRequest,
    *,
    existing_therapists: Iterable[Therapist] = (),
    existing_therapist_names: Iterable[str] = (),
) -> RegistrationCheck:
    issues: list[RegistrationIssue] = []
    name = request.display_name.strip()

    if not name:
        issues.append(RegistrationIssue("therapist_name_required", "display_name", "Therapist name is required"))
    else:
        existing = {_key(item.display_name) for item in existing_therapists}
        existing.update(_key(value) for value in existing_therapist_names)
        if _key(name) in existing:
            issues.append(RegistrationIssue("duplicate_therapist_name", "display_name", "Therapist already exists"))

    return RegistrationCheck(tuple(issues))


def validate_new_student(
    request: NewStudentRequest,
    *,
    existing_students: Iterable[Student] = (),
    known_therapist_ids: Iterable[str] | None = None,
) -> RegistrationCheck:
    issues: list[RegistrationIssue] = []
    students = tuple(existing_students)

    if not _nonblank(request.student_id):
        issues.append(RegistrationIssue("student_id_required", "student_id", "Student ID is required"))
    elif _key(request.student_id) in {_key(item.student_id) for item in students}:
        issues.append(RegistrationIssue("duplicate_student_id", "student_id", "Student ID already exists"))

    if not _nonblank(request.display_name):
        issues.append(RegistrationIssue("student_name_required", "display_name", "Student name is required"))

    if request.student_number < 1:
        issues.append(RegistrationIssue("invalid_student_number", "student_number", "Student number must be positive"))
    elif request.student_number in {item.student_number for item in students}:
        issues.append(RegistrationIssue("duplicate_student_number", "student_number", "Student number already exists"))

    if request.placement_end < request.placement_start:
        issues.append(RegistrationIssue("invalid_placement_dates", "placement_end", "Placement end cannot be before placement start"))

    if request.supervisor_therapist_id is not None and known_therapist_ids is not None:
        known = {_key(value) for value in known_therapist_ids}
        if _key(request.supervisor_therapist_id) not in known:
            issues.append(RegistrationIssue("unknown_supervisor", "supervisor_therapist_id", "Supervisor therapist is not registered"))

    return RegistrationCheck(tuple(issues))
