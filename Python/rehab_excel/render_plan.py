from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from enum import Enum
from pathlib import Path
from typing import Iterable, Mapping
import unicodedata

from rehab_core.daily_state import DailySessionState, DailySessionStatus
from rehab_core.models import Patient, ReplacementProviderKind, Session, Student
from rehab_core.students import student_display_state

from .layout_contract import (
    LayoutSafetyError,
    resolve_master_schedule_rows,
    resolve_therapist_daily_cell,
)


class RenderLineRole(str, Enum):
    ACTIVE = "active"
    ORIGINAL = "original"
    REPLACEMENT = "replacement"
    ABSENCE = "absence"


class RenderFontRole(str, Enum):
    DEFAULT = "default"
    STUDENT_ACTIVE = "student_active_green"
    MUTED = "muted"


class RenderFillRole(str, Enum):
    DEFAULT = "default"
    INFECTIOUS = "infectious_yellow"
    ROBOTIC = "robotic_pink"


class RenderBorderRole(str, Enum):
    DEFAULT = "default"
    INFECTIOUS = "infectious_yellow"


@dataclass(frozen=True)
class RenderLine:
    text: str
    role: RenderLineRole
    strike_through: bool = False
    font_role: RenderFontRole = RenderFontRole.DEFAULT


@dataclass(frozen=True)
class CellRenderPlan:
    sheet: str
    cell: str
    lines: tuple[RenderLine, ...]
    fill_role: RenderFillRole = RenderFillRole.DEFAULT
    border_role: RenderBorderRole = RenderBorderRole.DEFAULT
    wrap_text: bool = True
    min_font_size: int = 10

    @property
    def text(self) -> str:
        return "\n".join(line.text for line in self.lines)


@dataclass(frozen=True)
class SessionCellBinding:
    session_id: str
    patient_id: str
    master_schedule_cell: str
    original_daily_cell: str
    effective_daily_cell: str | None


@dataclass(frozen=True)
class RenderIssue:
    code: str
    message: str
    severity: str = "error"


@dataclass(frozen=True)
class DailyExcelRenderPlan:
    workbook_path: str
    target_date: date
    bindings: tuple[SessionCellBinding, ...]
    cells: tuple[CellRenderPlan, ...]
    issues: tuple[RenderIssue, ...]

    @property
    def ok(self) -> bool:
        return not any(issue.severity == "error" for issue in self.issues)


_TREATMENT_COLUMNS: Mapping[str, str] = {
    "ΦΘ": "D",
    "ΡΟΜΠΟΤΙΚΟ": "E",
    "ΠΙΣΙΝΑ": "F",
    "ΑΝΑΚΛΙΝΟΜΕΝΟ": "G",
    "ΕΡΓΟ": "H",
    "ΛΟΓΟ": "I",
    "ΕΦΑ": "J",
}


def _treatment_key(value: str) -> str:
    text = str(value).strip().upper()
    decomposed = unicodedata.normalize("NFD", text)
    return "".join(char for char in decomposed if not unicodedata.combining(char))


def _treatment_column(treatment: str | None) -> str | None:
    if treatment is None:
        return None
    return _TREATMENT_COLUMNS.get(_treatment_key(treatment))


def _provider_label(provider_id: str, labels: Mapping[str, str]) -> str:
    return labels.get(provider_id, provider_id)


def _student_lookup_aliases(student: Student, target_date: date) -> tuple[str, ...]:
    display = student_display_state(student, target_date)
    # Current v27 uses labels such as "φοιτ 1". The future UI will show the
    # real active name, but the mapper must still find the existing column.
    return (
        display.label,
        f"φοιτ {student.student_number}",
        f"φοιτητής {student.student_number}",
    )


def _replacement_provider_display(
    provider_id: str,
    *,
    provider_kind: ReplacementProviderKind,
    labels: Mapping[str, str],
    students: Mapping[str, Student],
    target_date: date,
) -> tuple[str, RenderFontRole]:
    """Return the human label used on the original slot's replacement line."""

    if provider_kind == ReplacementProviderKind.STUDENT:
        student = students.get(provider_id)
        if student is not None:
            display = student_display_state(student, target_date)
            return (
                display.label,
                RenderFontRole.STUDENT_ACTIVE
                if display.use_green_font
                else RenderFontRole.DEFAULT,
            )
    return _provider_label(provider_id, labels), RenderFontRole.DEFAULT


def _resolve_provider_cell(
    workbook_path: str | Path,
    provider_id: str,
    slot_time,
    *,
    labels: Mapping[str, str],
    provider_kind: ReplacementProviderKind,
    students: Mapping[str, Student],
    target_date: date,
) -> str:
    candidates: list[str] = []
    if provider_kind == ReplacementProviderKind.STUDENT:
        student = students.get(provider_id)
        if student is not None:
            candidates.extend(_student_lookup_aliases(student, target_date))
    candidates.append(_provider_label(provider_id, labels))
    candidates.append(provider_id)

    seen: set[str] = set()
    last_error: Exception | None = None
    for candidate in candidates:
        if not candidate or candidate in seen:
            continue
        seen.add(candidate)
        try:
            return resolve_therapist_daily_cell(workbook_path, candidate, slot_time)
        except LayoutSafetyError as exc:
            last_error = exc
    if last_error is not None:
        raise last_error
    raise LayoutSafetyError(
        f"No THERAPIST_DAILY provider label available for {provider_id!r}"
    )


def _merge_fill(current: RenderFillRole, *, robotic: bool, infectious: bool) -> RenderFillRole:
    # Robotic remains pink. Infectious is still preserved independently by a
    # yellow border so both signals remain visible on a robotic infectious case.
    if robotic:
        return RenderFillRole.ROBOTIC
    if infectious:
        return RenderFillRole.INFECTIOUS
    return current


def _merge_border(current: RenderBorderRole, *, infectious: bool) -> RenderBorderRole:
    if infectious:
        return RenderBorderRole.INFECTIOUS
    return current


def build_daily_excel_render_plan(
    workbook_path: str | Path,
    states: Iterable[DailySessionState],
    sessions: Iterable[Session],
    patients: Iterable[Patient],
    *,
    provider_labels: Mapping[str, str] | None = None,
    students: Iterable[Student] = (),
) -> DailyExcelRenderPlan:
    """Map operational daily state to verified workbook cells without writing.

    The plan deliberately renders only THERAPIST_DAILY. MASTER_SCHEDULE cells
    are resolved as references in ``SessionCellBinding`` but are not mutated by
    this stage, because MASTER_SCHEDULE contains recurring programme text and a
    one-day overlay must not erase day-pattern information.

    A later Excel writer may consume these semantic formatting roles. This
    function itself never opens the workbook in write mode and never saves it.
    """

    state_list = tuple(states)
    session_by_id = {session.session_id: session for session in sessions}
    patient_by_id = {patient.patient_id: patient for patient in patients}
    student_by_id = {student.student_id: student for student in students}
    labels = dict(provider_labels or {})

    dates = {state.session_date for state in state_list}
    if len(dates) > 1:
        raise ValueError("Daily Excel render plan requires states from one date only")
    target_date = next(iter(dates), date.today())

    issues: list[RenderIssue] = []
    bindings: list[SessionCellBinding] = []

    # Build per-cell line contributions first because one provider/time cell can
    # legitimately contain more than one patient in the legacy schedule.
    cell_lines: dict[str, list[RenderLine]] = {}
    cell_fill: dict[str, RenderFillRole] = {}
    cell_border: dict[str, RenderBorderRole] = {}

    for state in state_list:
        session = session_by_id.get(state.session_id)
        if session is None:
            issues.append(
                RenderIssue(
                    "MISSING_SESSION",
                    f"No Session found for DailySessionState {state.session_id}",
                )
            )
            continue
        patient = patient_by_id.get(state.patient_id)
        if patient is None:
            issues.append(
                RenderIssue(
                    "MISSING_PATIENT",
                    f"No Patient found for id {state.patient_id} ({state.session_id})",
                )
            )
            continue

        treatment_col = _treatment_column(session.treatment)
        if treatment_col is None:
            issues.append(
                RenderIssue(
                    "UNSUPPORTED_TREATMENT",
                    f"Unsupported treatment {session.treatment!r} for {state.session_id}",
                )
            )
            continue

        rows = resolve_master_schedule_rows(workbook_path, patient.display_name)
        if len(rows) != 1:
            issues.append(
                RenderIssue(
                    "MASTER_PATIENT_ROW_AMBIGUOUS",
                    f"Expected one MASTER_SCHEDULE row for {patient.display_name!r}; found {rows}",
                )
            )
            continue
        master_cell = f"{treatment_col}{rows[0]}"

        try:
            original_daily_cell = _resolve_provider_cell(
                workbook_path,
                state.original_therapist_id,
                state.original_time,
                labels=labels,
                provider_kind=ReplacementProviderKind.THERAPIST,
                students=student_by_id,
                target_date=target_date,
            )
        except LayoutSafetyError as exc:
            issues.append(
                RenderIssue(
                    "ORIGINAL_DAILY_CELL_NOT_FOUND",
                    f"{state.session_id}: {exc}",
                )
            )
            continue

        effective_daily_cell: str | None = None
        if (
            state.status == DailySessionStatus.REPLACED
            and state.effective_therapist_id is not None
            and state.effective_time is not None
        ):
            try:
                effective_daily_cell = _resolve_provider_cell(
                    workbook_path,
                    state.effective_therapist_id,
                    state.effective_time,
                    labels=labels,
                    provider_kind=(
                        state.effective_provider_kind
                        or ReplacementProviderKind.THERAPIST
                    ),
                    students=student_by_id,
                    target_date=target_date,
                )
            except LayoutSafetyError as exc:
                issues.append(
                    RenderIssue(
                        "EFFECTIVE_DAILY_CELL_NOT_FOUND",
                        f"{state.session_id}: {exc}",
                    )
                )

        bindings.append(
            SessionCellBinding(
                session_id=state.session_id,
                patient_id=state.patient_id,
                master_schedule_cell=master_cell,
                original_daily_cell=original_daily_cell,
                effective_daily_cell=effective_daily_cell,
            )
        )

        original_strike = state.status in {
            DailySessionStatus.PATIENT_ABSENT,
            DailySessionStatus.THERAPIST_ABSENT,
            DailySessionStatus.REPLACED,
        }
        original_role = (
            RenderLineRole.ACTIVE
            if state.status == DailySessionStatus.ACTIVE
            else RenderLineRole.ORIGINAL
        )
        original_font = (
            RenderFontRole.DEFAULT
            if state.status == DailySessionStatus.ACTIVE
            else RenderFontRole.MUTED
        )
        cell_lines.setdefault(original_daily_cell, []).append(
            RenderLine(
                patient.display_name,
                role=original_role,
                strike_through=original_strike,
                font_role=original_font,
            )
        )
        cell_fill[original_daily_cell] = _merge_fill(
            cell_fill.get(original_daily_cell, RenderFillRole.DEFAULT),
            robotic=session.robotic,
            infectious=patient.infectious,
        )
        cell_border[original_daily_cell] = _merge_border(
            cell_border.get(original_daily_cell, RenderBorderRole.DEFAULT),
            infectious=patient.infectious,
        )

        if state.status == DailySessionStatus.REPLACED and effective_daily_cell is not None:
            provider_kind = (
                state.effective_provider_kind or ReplacementProviderKind.THERAPIST
            )
            provider_label, provider_font = _replacement_provider_display(
                state.effective_therapist_id,
                provider_kind=provider_kind,
                labels=labels,
                students=student_by_id,
                target_date=target_date,
            )
            # Keep the original slot visually self-contained: the original
            # patient line is struck through and the effective provider/time is
            # shown immediately below it. This preserves the compact timetable
            # width while making the daily exception readable at a glance.
            cell_lines.setdefault(original_daily_cell, []).append(
                RenderLine(
                    f"→ {provider_label} {state.effective_time.strftime('%H:%M')}",
                    role=RenderLineRole.REPLACEMENT,
                    strike_through=False,
                    font_role=provider_font,
                )
            )

            if effective_daily_cell != original_daily_cell:
                cell_lines.setdefault(effective_daily_cell, []).append(
                    RenderLine(
                        patient.display_name,
                        role=RenderLineRole.REPLACEMENT,
                        strike_through=False,
                        font_role=RenderFontRole.DEFAULT,
                    )
                )
                cell_fill[effective_daily_cell] = _merge_fill(
                    cell_fill.get(effective_daily_cell, RenderFillRole.DEFAULT),
                    robotic=session.robotic,
                    infectious=patient.infectious,
                )
                cell_border[effective_daily_cell] = _merge_border(
                    cell_border.get(effective_daily_cell, RenderBorderRole.DEFAULT),
                    infectious=patient.infectious,
                )

    cells = tuple(
        CellRenderPlan(
            sheet="THERAPIST_DAILY",
            cell=cell,
            lines=tuple(lines),
            fill_role=cell_fill.get(cell, RenderFillRole.DEFAULT),
            border_role=cell_border.get(cell, RenderBorderRole.DEFAULT),
        )
        for cell, lines in sorted(
            cell_lines.items(), key=lambda item: (int(''.join(filter(str.isdigit, item[0]))), item[0])
        )
    )

    return DailyExcelRenderPlan(
        workbook_path=str(workbook_path),
        target_date=target_date,
        bindings=tuple(bindings),
        cells=cells,
        issues=tuple(issues),
    )
