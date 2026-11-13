"""Read-only, source-bound planning contract for a private real-source trial.

This is NOT a workbook importer, writer, approval or release. It never logs
patient information; callers must keep the returned in-memory plan private.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from hashlib import sha256
from pathlib import Path

from rehab_core.programme_mapping import SourceWeeklyEntry
from rehab_excel.trial_admission_gate import evaluate_trial_admission
from tools.preflight_trial_excel_sources import inspect_native_pilot_inputs


class TrialImportPlanError(ValueError):
    """A trial plan cannot be created or a source is stale."""


def _digest(path: Path) -> str:
    h = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


@dataclass(frozen=True)
class TrialImportPlan:
    """Private metadata and exact approved weekday placements, no writes."""

    review_sha256: str
    intake_sha256: str
    patient_count: int
    weekday_slot_count: int
    entries: tuple[SourceWeeklyEntry, ...] = field(repr=False)

    def verify_sources(self, review: str | Path, intake: str | Path) -> None:
        """Require both original files unchanged immediately before future use."""
        r, c = Path(review).resolve(), Path(intake).resolve()
        if not r.is_file() or not c.is_file():
            raise TrialImportPlanError("Trial input missing")
        if _digest(r) != self.review_sha256 or _digest(c) != self.intake_sha256:
            raise TrialImportPlanError("SAFETY STOP: stale trial input")


def prepare_trial_import_plan(
    *, review: str | Path, intake: str | Path
) -> TrialImportPlan:
    """Validate both private inputs and preserve the exact admitted placements.

    No partial imports, identities inferred from names, old synthetic history,
    or workbook modifications are allowed.
    """
    r, c = Path(review).resolve(), Path(intake).resolve()
    if r == c or r.suffix.lower() != ".xlsx" or c.suffix.lower() != ".xlsm":
        raise TrialImportPlanError("Separate review XLSX and pilot XLSM required")
    if not r.is_file() or not c.is_file():
        raise TrialImportPlanError("Trial input missing")
    initial = (_digest(r), _digest(c))
    try:
        status = inspect_native_pilot_inputs(review=r, cc=c)
        if not status["native_trial_permitted"]:
            raise TrialImportPlanError("SAFETY STOP: trial admission is blocked")
        admissions = evaluate_trial_admission(r)
        if not admissions.ready or not admissions.admitted_entries:
            raise TrialImportPlanError("SAFETY STOP: no fully approved schedule")
        if (
            admissions.candidate_patients != status["review_patients"]
            or admissions.verified_patient_ids != status["verified_patient_ids"]
            or len(admissions.admitted_entries) != status["proposed_weekly_slots"]
        ):
            raise TrialImportPlanError("SAFETY STOP: staging results changed")
        plan = TrialImportPlan(
            review_sha256=initial[0],
            intake_sha256=initial[1],
            patient_count=admissions.verified_patient_ids,
            weekday_slot_count=len(admissions.admitted_entries),
            entries=tuple(admissions.admitted_entries),
        )
    finally:
        if (_digest(r), _digest(c)) != initial:
            raise TrialImportPlanError("SAFETY STOP: source changed during planning")
    return plan
