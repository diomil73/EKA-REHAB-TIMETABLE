from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
from pathlib import Path
from typing import Iterable, Mapping

from rehab_core.daily_state import DailySessionState, DailySessionStatus
from rehab_core.models import (
    BaseScheduleEntry,
    Patient,
    ReplacementProviderKind,
    Student,
)
from rehab_core.students import student_display_state

from .layout_contract import LayoutSafetyError, resolve_therapist_daily_cell
from .writeback import CellPatch, WriteIntent


class OutpatientPresentationError(RuntimeError):
    """Raised when outpatient presentation would require an unsafe guess."""


@dataclass(frozen=True)
class OutpatientDailyTarget:
    """One effective THERAPIST_DAILY slot that must be outpatient blue."""

    provider_id: str
    provider_kind: ReplacementProviderKind
    start_time: time
    patient_ids: tuple[str, ...]


def _is_robotic_treatment(value: str) -> bool:
    normalized = value.strip().casefold()
    return "ρομποτ" in normalized or "robot" in normalized


def validate_outpatient_base_entries(
    entries: Iterable[BaseScheduleEntry],
    patients: Iterable[Patient],
) -> None:
    """Reject robotic recurring assignments for outpatients.

    Confirmed business rule: an outpatient may participate in the normal
    therapist programme, but may never be assigned robotic treatment as an
    extra/robotic programme. The rule is checked both from the explicit
    ``robotic`` flag and from the treatment label so an inconsistent import
    cannot silently bypass it.
    """

    patient_by_id = {patient.patient_id: patient for patient in patients}
    for entry in entries:
        patient = patient_by_id.get(entry.patient_id)
        if patient is None or not patient.is_outpatient:
            continue
        if entry.robotic or _is_robotic_treatment(entry.treatment):
            raise OutpatientPresentationError(
                f"Outpatient {entry.patient_id!r} cannot have robotic treatment"
            )


def _effective_target(
    state: DailySessionState,
) -> tuple[ReplacementProviderKind, str, time] | None:
    if state.status == DailySessionStatus.ACTIVE:
        return (
            ReplacementProviderKind.THERAPIST,
            state.original_therapist_id,
            state.original_time,
        )
    if state.status == DailySessionStatus.REPLACED:
        if state.effective_therapist_id is None or state.effective_time is None:
            raise OutpatientPresentationError(
                f"Replacement state {state.session_id!r} lacks effective provider/time"
            )
        return (
            state.effective_provider_kind or ReplacementProviderKind.THERAPIST,
            state.effective_therapist_id,
            state.effective_time,
        )
    return None


def outpatient_daily_targets(
    states: Iterable[DailySessionState],
    patients: Iterable[Patient],
) -> tuple[OutpatientDailyTarget, ...]:
    """Return effective daily slots that must be light blue.

    Colour is determined from the patient(s) actually active in the slot on the
    concrete day, not from every patient who may share the same recurring slot
    on other weekdays. If one effective daily cell would contain both an
    inpatient and an outpatient at the same time, the renderer refuses to guess
    one cell-level fill colour and raises instead.
    """

    patient_by_id = {patient.patient_id: patient for patient in patients}
    occupants: dict[
        tuple[ReplacementProviderKind, str, time], list[Patient]
    ] = {}

    for state in states:
        target = _effective_target(state)
        if target is None:
            continue
        patient = patient_by_id.get(state.patient_id)
        if patient is None:
            raise OutpatientPresentationError(
                f"Missing patient {state.patient_id!r} for daily presentation"
            )
        occupants.setdefault(target, []).append(patient)

    targets: list[OutpatientDailyTarget] = []
    for (provider_kind, provider_id, start_time), slot_patients in occupants.items():
        outpatient_flags = {patient.is_outpatient for patient in slot_patients}
        if len(outpatient_flags) > 1:
            raise OutpatientPresentationError(
                f"Cannot apply one cell colour to mixed inpatient/outpatient slot "
                f"{provider_id!r} at {start_time.strftime('%H:%M')}"
            )
        if True not in outpatient_flags:
            continue
        targets.append(
            OutpatientDailyTarget(
                provider_id=provider_id,
                provider_kind=provider_kind,
                start_time=start_time,
                patient_ids=tuple(patient.patient_id for patient in slot_patients),
            )
        )

    targets.sort(
        key=lambda target: (
            target.start_time,
            target.provider_kind.value,
            target.provider_id.casefold(),
        )
    )
    return tuple(targets)


def _provider_cell(
    workbook_path: Path,
    target: OutpatientDailyTarget,
    *,
    provider_labels: Mapping[str, str],
    students: Mapping[str, Student],
    target_date: date,
) -> str:
    candidates: list[str] = []
    if target.provider_kind == ReplacementProviderKind.STUDENT:
        student = students.get(target.provider_id)
        if student is not None:
            state = student_display_state(student, target_date)
            candidates.extend(
                (
                    state.label,
                    f"φοιτ {student.student_number}",
                    f"φοιτητής {student.student_number}",
                )
            )
    candidates.extend(
        (
            provider_labels.get(target.provider_id, target.provider_id),
            target.provider_id,
        )
    )

    last_error: Exception | None = None
    seen: set[str] = set()
    for candidate in candidates:
        if not candidate or candidate in seen:
            continue
        seen.add(candidate)
        try:
            return resolve_therapist_daily_cell(
                workbook_path,
                candidate,
                target.start_time,
            )
        except LayoutSafetyError as exc:
            last_error = exc
    raise OutpatientPresentationError(
        str(last_error or f"No THERAPIST_DAILY cell for provider {target.provider_id!r}")
    )


def build_outpatient_daily_patches(
    workbook_path: str | Path,
    *,
    states: Iterable[DailySessionState],
    patients: Iterable[Patient],
    provider_labels: Mapping[str, str] | None = None,
    students: Iterable[Student] = (),
) -> tuple[CellPatch, ...]:
    """Build presentation-only blue-fill patches for active outpatient slots."""

    state_list = tuple(states)
    dates = {state.session_date for state in state_list}
    if len(dates) != 1:
        raise OutpatientPresentationError(
            "Outpatient daily presentation requires states from exactly one date"
        )
    target_date = next(iter(dates))
    labels = dict(provider_labels or {})
    student_map = {student.student_id: student for student in students}
    path = Path(workbook_path)

    patches: list[CellPatch] = []
    for target in outpatient_daily_targets(state_list, patients):
        patches.append(
            CellPatch(
                sheet="THERAPIST_DAILY",
                cell=_provider_cell(
                    path,
                    target,
                    provider_labels=labels,
                    students=student_map,
                    target_date=target_date,
                ),
                intent=WriteIntent.PRESENTATION,
                fill_role="outpatient_light_blue",
                source_tag="outpatient_daily_presentation",
            )
        )
    return tuple(patches)
