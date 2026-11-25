"""Independent OOXML preflight assembly for one DAILY_INPUT projection date.

This assembles existing canonical readers; it never writes the workbook. The
native COM diagnostics run separately and bind to this exact source digest.
No phase report, receipt, trust attestation or Apply authorization is created.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from rehab_core.base_schedule import materialize_sessions_for_date
from rehab_core.daily_state import DailySessionState, build_daily_session_states
from rehab_core.lifecycle import active_entries_on
from rehab_core.models import (
    BaseScheduleEntry, Patient, ReplacementProviderKind, ReplacementAssignment,
    Session, Student,
)
from rehab_core.recurring_therapy_changes import (
    RecurringTherapyChange, recurring_entries_on_date,
    validate_effective_weekly_patient_slots,
)

from .authoritative_commit import file_sha256
from .daily_input_reader import DailyInputReadResult, read_daily_input
from .daily_replacement_facts import read_daily_replacement_facts
from .daily_input_therapist_projection_expected import (
    TherapistDailyExpectedGrid, build_therapist_daily_expected_grid,
)
from .master_projection import MasterProjectionRow, build_master_projection
from .outpatient_schedule_source import read_unified_base_schedule
from .package_probe import WorkbookPackageProbe
from .patient_exit_source import read_patient_exits
from .patient_registry_source import read_patient_registry
from .reader import WorkbookSettings, read_settings
from .recurring_change_event_source import read_recurring_change_events
from .stable_recurring_ids import bind_stable_assignment_ids, is_stable_assignment_id
from .student_registry import read_students
from .daily_input_therapist_projection_readback import _name


class DailyInputProjectionPreflightError(RuntimeError):
    """The independent source preflight is incomplete or inconsistent."""


@dataclass(frozen=True)
class DailyInputProjectionPreflight:
    service_date: date
    baseline_date: date
    source_workbook_sha256: str
    patients: tuple[Patient, ...] = field(repr=False)
    settings: WorkbookSettings = field(repr=False)
    stable_base_entries: tuple[BaseScheduleEntry, ...] = field(repr=False)
    effective_base_entries: tuple[BaseScheduleEntry, ...] = field(repr=False)
    recurring_changes: tuple[RecurringTherapyChange, ...] = field(repr=False)
    patient_exits: tuple = field(repr=False)
    sessions: tuple[Session, ...] = field(repr=False)
    daily_input: DailyInputReadResult = field(repr=False)
    replacements: tuple[ReplacementAssignment, ...] = field(repr=False)
    states: tuple[DailySessionState, ...] = field(repr=False)
    students: tuple[Student, ...] = field(repr=False)
    expected_projection: tuple[MasterProjectionRow, ...] = field(repr=False)
    expected_therapist_grid: TherapistDailyExpectedGrid = field(repr=False)

    @property
    def validated_changes(self) -> tuple:
        return self.daily_input.source_changes

    @property
    def patient_count(self) -> int:
        return len(self.patients)

    @property
    def base_assignment_count(self) -> int:
        return len(self.stable_base_entries)

    @property
    def session_count(self) -> int:
        return len(self.sessions)

    @property
    def daily_change_count(self) -> int:
        return len(self.daily_input.source_changes)

    @property
    def replacement_count(self) -> int:
        return len(self.replacements)

    @property
    def therapist_assignment_count(self) -> int:
        return len(self.expected_therapist_grid.expected_assignments)


def _provider_headers(path: Path) -> tuple[str, ...]:
    try:
        cells = WorkbookPackageProbe(path).cell_values("THERAPIST_DAILY")
        return tuple(
            str(cells.get(f"{chr(ord('B') + col)}{row}") or "").strip()
            for row in (1, 11)
            for col in range(9)
        )
    except Exception as exc:
        raise DailyInputProjectionPreflightError("THERAPIST_DAILY provider layout unavailable") from exc


def _provider_roster(settings: WorkbookSettings, students: tuple[Student, ...]) -> tuple[set[str], dict[str, str]]:
    labels = (
        *settings.therapist_names,
        *settings.psychologist_names,
        *settings.occupational_therapist_names,
        *settings.speech_therapist_names,
    )
    allowed = { _name(value) for value in labels if str(value).strip() }
    for student in students:
        allowed.update({
            _name(student.display_name),
            _name(student.student_id),
            _name(f"φοιτ {student.student_id}"),
            _name(f"φοιτητής {student.student_id}"),
        })
    provider_labels = {str(value): str(value) for value in labels if str(value).strip()}
    return allowed, provider_labels


def build_daily_input_projection_preflight(
    workbook_path: str | Path,
    *,
    service_date: date,
    baseline_date: date,
) -> DailyInputProjectionPreflight:
    """Assemble canonical source facts for native read-only verification.

    baseline_date is explicit because recurring change history is only valid
    relative to a known programme baseline; it is never inferred from workbook
    filenames or the current date. The resulting facts remain non-authorizing.
    """
    if type(service_date) is not date or type(baseline_date) is not date:
        raise DailyInputProjectionPreflightError("Explicit service and programme baseline dates required")
    if service_date < baseline_date:
        raise DailyInputProjectionPreflightError("Service date predates programme baseline")

    try:
        source = Path(workbook_path).resolve(strict=True)
        if not source.is_file() or source.suffix.casefold() != ".xlsm":
            raise DailyInputProjectionPreflightError("Existing saved XLSM required")
        before_hash = file_sha256(source)

        patients = tuple(read_patient_registry(source))
        patient_ids = [patient.patient_id.casefold() for patient in patients]
        if not patients or len(set(patient_ids)) != len(patient_ids):
            raise DailyInputProjectionPreflightError("Patient registry is empty or has duplicate identities")

        settings = read_settings(source)
        if not settings.therapist_names:
            raise DailyInputProjectionPreflightError("Therapist roster is empty")

        raw_entries = tuple(read_unified_base_schedule(source, patients=patients))
        if not raw_entries:
            raise DailyInputProjectionPreflightError("Unified recurring programme is empty")
        stable_entries = tuple(bind_stable_assignment_ids(source, raw_entries, require_complete=True))
        if (len(stable_entries) != len(raw_entries)
                or any(not is_stable_assignment_id(item.base_entry_id) for item in stable_entries)
                or len({item.base_entry_id for item in stable_entries}) != len(stable_entries)):
            raise DailyInputProjectionPreflightError("Recurring assignment identity is incomplete or duplicated")

        recurring_changes = tuple(read_recurring_change_events(
            source, entries=stable_entries, baseline_date=baseline_date,
        ))
        effective = tuple(recurring_entries_on_date(
            stable_entries,
            baseline_date=baseline_date,
            target_date=service_date,
            changes=recurring_changes,
        ))
        exits = tuple(read_patient_exits(source, patients=patients))
        effective = tuple(active_entries_on(effective, service_date, exits))
        if (any(not is_stable_assignment_id(item.base_entry_id) for item in effective)
                or len({item.base_entry_id for item in effective}) != len(effective)):
            raise DailyInputProjectionPreflightError("Effective recurring assignment identity is invalid")
        known_patient_ids = {patient.patient_id for patient in patients}
        if any(item.patient_id not in known_patient_ids for item in effective):
            raise DailyInputProjectionPreflightError("Effective recurring entry references an unknown patient")
        validate_effective_weekly_patient_slots(effective)

        sessions = tuple(materialize_sessions_for_date(effective, service_date))
        if not sessions:
            raise DailyInputProjectionPreflightError("No provider-owned sessions exist for this service date")

        daily_input = read_daily_input(
            source,
            patients=patients,
            sessions=sessions,
            therapist_names=settings.therapist_names,
            allow_empty=False,
        )
        if daily_input.target_date != service_date or not daily_input.source_changes:
            raise DailyInputProjectionPreflightError("DAILY_INPUT date or validated change set is invalid")

        replacements = tuple(read_daily_replacement_facts(
            source, sessions=sessions, target_date=service_date,
        ))
        states = tuple(build_daily_session_states(
            sessions,
            absences=daily_input.absences,
            replacements=replacements,
            cancellations=daily_input.cancellations,
            target_date=service_date,
        ))
        if len(states) != len(sessions) or {s.session_id for s in states} != {s.session_id for s in sessions}:
            raise DailyInputProjectionPreflightError("Daily session state does not cover exact dated session set")

        students = tuple(read_students(source))
        allowed_providers, provider_labels = _provider_roster(settings, students)
        provider_headers = _provider_headers(source)
        for header in provider_headers:
            if header and _name(header) not in allowed_providers:
                raise DailyInputProjectionPreflightError("THERAPIST_DAILY header is outside the approved provider roster")
        for entry in effective:
            if entry.therapist_id is not None and _name(entry.therapist_id) not in allowed_providers:
                raise DailyInputProjectionPreflightError("Recurring entry references a provider outside the approved roster")
        for replacement in replacements:
            if replacement.replacement_provider_kind == ReplacementProviderKind.THERAPIST:
                if _name(replacement.replacement_therapist_id) not in allowed_providers:
                    raise DailyInputProjectionPreflightError("Replacement references an unknown therapist")
            elif replacement.replacement_provider_kind == ReplacementProviderKind.STUDENT:
                if replacement.replacement_therapist_id not in {s.student_id for s in students}:
                    raise DailyInputProjectionPreflightError("Replacement references an unknown student")

        expected_grid = build_therapist_daily_expected_grid(
            source,
            service_date=service_date,
            expected_provider_headers=provider_headers,
            base_entries=effective,
            sessions=sessions,
            states=states,
            patients=patients,
            provider_labels=provider_labels,
            students=students,
        )
        expected_projection = tuple(build_master_projection(
            source, patients=patients, target_date=service_date,
        ))
        if file_sha256(source) != before_hash:
            raise DailyInputProjectionPreflightError("Workbook bytes changed during source preflight")

        return DailyInputProjectionPreflight(
            service_date=service_date,
            baseline_date=baseline_date,
            source_workbook_sha256=before_hash,
            patients=patients,
            settings=settings,
            stable_base_entries=stable_entries,
            effective_base_entries=effective,
            recurring_changes=recurring_changes,
            patient_exits=exits,
            sessions=sessions,
            daily_input=daily_input,
            replacements=replacements,
            states=states,
            students=students,
            expected_projection=expected_projection,
            expected_therapist_grid=expected_grid,
        )
    except DailyInputProjectionPreflightError:
        raise
    except Exception as exc:
        raise DailyInputProjectionPreflightError("Independent DAILY_INPUT projection preflight failed closed") from exc
