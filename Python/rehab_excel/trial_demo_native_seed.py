"""Native-only, isolated seed of one clean practice-demo workbook.

No output is written until original A:V data, review confirmations and weekly
placements have passed the full source-bound gate. Existing patients, schedule
rows, operational history or an occupied output path STOP the operation.

Only native Microsoft Excel COM may save an actual .xlsm. The final destination
is published with exclusive-create after independent readback, never overwriting
a frozen proof artifact or an existing DEMO. The workbook is still a practice
candidate until a human checks the interactive workflows.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from collections import Counter
from hashlib import sha256
from pathlib import Path
import shutil
import sys
import tempfile
from typing import Mapping, Protocol
from zipfile import ZipFile

from openpyxl import load_workbook

from rehab_core.programme_mapping import parse_clock
from rehab_excel.application_contract import assert_application_contract
from rehab_excel.master_projection_writer import refresh_master_projection_in_place
from rehab_excel.master_visual_style import apply_master_visual_style_in_place
from rehab_excel.patient_planner_projection import planner_identity_formulas
from rehab_excel.trial_demo_planner_mirror import (
    PLANNER_TREATMENT_HEADERS, choose_planner_demo_mirrors,
)
from rehab_excel.outpatient_schedule_source import read_unified_base_schedule
from rehab_excel.patient_registry_source import read_patient_registry
from rehab_excel.trial_demo_source_projection import (
    TrialDemoSourceProjection, project_trial_demo_sources,
)
from rehab_excel.trial_import_plan import TrialImportPlan
from rehab_excel.trial_patient_registry_projection import ConfirmedClinicalFields
from rehab_excel.weekly_slots_source import (
    HEADERS as WEEKLY_HEADERS, WEEKDAY_PATTERN,
)
from tools.audit_pilot_source_readiness import inspect_source


class TrialDemoNativeSeedError(RuntimeError):
    """The native pilot cannot be published safely."""


@dataclass(frozen=True)
class TrialDemoNativeSeedReport:
    patient_count: int
    weekday_slot_count: int
    source_unchanged: bool
    verified_native_readback: bool
    output_path: str = field(repr=False)


class TrialDemoSeedBackend(Protocol):
    def write(
        self, workbook_path: Path, projection: TrialDemoSourceProjection,
    ) -> None: ...


def _digest(path: Path) -> str:
    h = sha256()
    with path.open("rb") as stream:
        for piece in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(piece)
    return h.hexdigest()


def _nonempty(values) -> bool:
    return any(value is not None and str(value).strip() for value in values)


def _private_output_target(output: Path) -> bool:
    # A real-data DEMO must not be saved within any Git checkout.
    return not any((directory / ".git").exists() for directory in
                   (output.parent, *output.parent.parents))


def assert_clean_native_demo_shell(intake: str | Path) -> None:
    """Inspect the source WITHOUT saving it; refuse any existing patient data.

    Nonempty Excel template formulas in PATIENT_PLANNER A:E are permitted,
    but any pre-existing treatment in F onward is not. Other live sheets are
    protected through the existing historical-source inventory.
    """
    source = Path(intake).resolve()
    if source.suffix.casefold() != ".xlsm" or not source.is_file():
        raise TrialDemoNativeSeedError("Existing macro-enabled demo shell required")
    try:
        status = inspect_source(source)
        if (
            set(status["issues"]) != {"EMPTY_OR_MISSING_PATIENT_REGISTRY"}
            or status["patient_registry_rows"] != 0
            or status["history_rows"] != 0
        ):
            raise TrialDemoNativeSeedError("Demo shell contains existing clinical state")
        with ZipFile(source) as archive:
            if "xl/vbaProject.bin" not in archive.namelist():
                raise TrialDemoNativeSeedError("Demo shell has no VBA project")
        wb = load_workbook(source, read_only=True, data_only=False, keep_vba=True)
        try:
            required = {
                "PATIENTS", "PATIENT_PLANNER", "MASTER_SCHEDULE",
                "DAILY_INPUT", "THERAPIST_DAILY", "SETTINGS",
            }
            if not required.issubset(wb.sheetnames):
                raise TrialDemoNativeSeedError("Demo shell is missing core sheets")
            patients = wb["PATIENTS"]
            if tuple(patients.cell(1, col).value for col in range(1, 6)) != (
                "PatientID", "Θάλαμος", "Ασθενής", "Λοιμώδης", "Κατάσταση",
            ):
                raise TrialDemoNativeSeedError("PATIENTS header contract mismatch")
            for values in patients.iter_rows(min_row=2, values_only=True):
                if _nonempty(values):
                    raise TrialDemoNativeSeedError("Demo shell PATIENTS not empty")
            planner = wb["PATIENT_PLANNER"]
            if (
                planner.cell(1, 1).value != "PatientID"
                or planner.cell(1, 3).value != "Ασθενής"
            ):
                raise TrialDemoNativeSeedError("PATIENT_PLANNER header contract mismatch")
            for row in planner.iter_rows(
                min_row=2, min_col=6, max_col=max(planner.max_column, 6),
                values_only=True,
            ):
                if _nonempty(row):
                    raise TrialDemoNativeSeedError("Demo shell contains old treatments")
            # A clean practice seed cannot inherit hidden/old operational
            # bookings, student assignments or previous daily decisions.
            old_operational_sources = (
                "OUTPATIENT_SCHEDULE", "SESSIONS", "REPLACEMENTS",
                "REPLACEMENT_LOG", "DAILY_LOG", "DAILY_CHANGE_LEDGER",
                "CONFLICT_LOG", "THERAPIST_ATTENDANCE", "STUDENTS",
            )
            for source_name in old_operational_sources:
                if source_name not in wb.sheetnames:
                    continue
                for values in wb[source_name].iter_rows(
                    min_row=2, values_only=True,
                ):
                    if _nonempty(values):
                        raise TrialDemoNativeSeedError(
                            "Demo shell contains old operational records"
                        )
            if "PROGRAM_WEEKLY_SLOTS" in wb.sheetnames:
                weekly = wb["PROGRAM_WEEKLY_SLOTS"]
                if tuple(weekly.cell(1, i).value for i in range(1, 7)) != WEEKLY_HEADERS:
                    raise TrialDemoNativeSeedError("Existing weekly sheet schema mismatch")
                for row in weekly.iter_rows(min_row=2, max_col=6, values_only=True):
                    if _nonempty(row):
                        raise TrialDemoNativeSeedError("Demo shell has existing weekly slots")
        finally:
            wb.close()
    except TrialDemoNativeSeedError:
        raise
    except Exception:
        raise TrialDemoNativeSeedError(
            "SAFETY STOP: unreadable or modified demo shell"
        ) from None


class Win32ComTrialDemoSeedBackend:
    """Use ONE Excel transaction to write initial PATIENTS + granular weekdays."""

    def write(
        self, workbook_path: Path, projection: TrialDemoSourceProjection,
    ) -> None:
        if sys.platform != "win32":
            raise TrialDemoNativeSeedError("Microsoft Excel for Windows required")
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError:
            raise TrialDemoNativeSeedError("pywin32 required for native Excel write") from None

        excel = None
        wb = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            wb = excel.Workbooks.Open(
                str(workbook_path.resolve()), UpdateLinks=0, ReadOnly=False,
            )
            patients = wb.Worksheets("PATIENTS")
            planner = wb.Worksheets("PATIENT_PLANNER")
            # Confirm the private temp copy is still pristine before any write.
            if patients.Cells(2, 1).Value2 not in (None, ""):
                raise TrialDemoNativeSeedError("Demo shell is no longer empty")

            headers = {}
            last_col = max(9, int(patients.UsedRange.Columns.Count))
            for col in range(1, last_col + 1):
                key = str(patients.Cells(1, col).Value2 or "").strip()
                if key:
                    headers[key] = col
            extra = ("PatientType", "HospitalMRN", "ResponsibleDoctor", "SourceContext")
            desired_cols = []
            next_col = max((5, *headers.values())) + 1
            for name in extra:
                if name in headers:
                    desired_cols.append(headers[name])
                    continue
                patients.Cells(1, next_col).Value = name
                desired_cols.append(next_col)
                next_col += 1
            for rownum, record in enumerate(projection.patient_registry.rows, start=2):
                if patients.Cells(rownum, 1).Value2 not in (None, ""):
                    raise TrialDemoNativeSeedError("Demo patient target row is occupied")
                patients.Cells(rownum, 1).NumberFormat = "@"
                for col in range(1, 6):
                    patients.Cells(rownum, col).Value = record[col - 1]
                for col, value in zip(desired_cols, record[5:]):
                    patients.Cells(rownum, col).Value = value
                for col, formula in enumerate(planner_identity_formulas(rownum), start=1):
                    planner.Cells(rownum, col).Formula = formula

            # Mirror simple, lossless recurring times into the familiar
            # PATIENT_PLANNER / MASTER UI, while detailed weekly slots remain
            # authoritative for the scheduling/replacement engine.
            planner_headers = {}
            for col in range(1, int(planner.UsedRange.Columns.Count) + 1):
                key = str(planner.Cells(1, col).Value2 or "").strip()
                if key:
                    planner_headers[key] = col
            planner_row_by_id = {
                str(item[0]): row for row, item in
                enumerate(projection.patient_registry.rows, start=2)
            }
            for mirror in choose_planner_demo_mirrors(projection.weekly_slots.rows):
                time_key, day_key, provider_key = PLANNER_TREATMENT_HEADERS[mirror.treatment]
                fields = (time_key, day_key) + ((provider_key,) if provider_key else ())
                if any(field not in planner_headers for field in fields):
                    raise TrialDemoNativeSeedError(
                        "Demo planner has missing treatment display columns"
                    )
                target_row = planner_row_by_id[mirror.patient_id]
                targets = [planner.Cells(target_row, planner_headers[field]) for field in fields]
                if any(cell.Value2 not in (None, "") for cell in targets):
                    raise TrialDemoNativeSeedError(
                        "Demo planner treatment target is already occupied"
                    )
                targets[0].NumberFormat = "hh:mm"
                targets[0].Value = (
                    mirror.start_time.hour * 3600 +
                    mirror.start_time.minute * 60
                ) / 86400
                targets[1].Value = mirror.day_pattern
                if provider_key:
                    targets[2].Value = mirror.therapist

            try:
                weekly = wb.Worksheets("PROGRAM_WEEKLY_SLOTS")
            except Exception:
                weekly = wb.Worksheets.Add(
                    After=wb.Worksheets(wb.Worksheets.Count)
                )
                weekly.Name = "PROGRAM_WEEKLY_SLOTS"
            for col, header in enumerate(WEEKLY_HEADERS, start=1):
                existing = str(weekly.Cells(1, col).Value2 or "").strip()
                if existing not in ("", header):
                    raise TrialDemoNativeSeedError("Weekly source header collision")
                weekly.Cells(1, col).Value = header
            for rownum, record in enumerate(projection.weekly_slots.rows, start=2):
                if weekly.Cells(rownum, 1).Value2 not in (None, ""):
                    raise TrialDemoNativeSeedError("Weekly target row occupied")
                for col, value in enumerate(record, start=1):
                    if col == 4:
                        value = (value.hour * 3600 + value.minute * 60) / 86400
                        weekly.Cells(rownum, col).NumberFormat = "hh:mm"
                    weekly.Cells(rownum, col).Value = value

            # Daily engine reads PROGRAM_WEEKLY_SLOTS, but the legacy MASTER
            # single-slot visual cells cannot losslessly display multiple
            # time/day variations. Provide an explicit uncompressed view.
            try:
                view = wb.Worksheets("DEMO_WEEKLY_VIEW")
            except Exception:
                view = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
                view.Name = "DEMO_WEEKLY_VIEW"
            view_headers = (
                "PatientID", "Ασθενής", "Θεραπεία", "Ημέρα",
                "Ώρα", "Θεραπευτής", "Κελί πηγής",
            )
            names = {
                str(row[0]): str(row[2])
                for row in projection.patient_registry.rows
            }
            for col, value in enumerate(view_headers, start=1):
                view.Cells(1, col).Value = value
            for rownum, record in enumerate(projection.weekly_slots.rows, start=2):
                pid, treatment, weekday, start, therapist, source_cell = record
                values = (
                    pid, names[str(pid)], treatment, weekday, start, therapist,
                    source_cell,
                )
                for col, value in enumerate(values, start=1):
                    if col == 5:
                        value = (start.hour * 3600 + start.minute * 60) / 86400
                        view.Cells(rownum, col).NumberFormat = "hh:mm"
                    view.Cells(rownum, col).Value = value
            # Preserve all existing VBA components and operational controls.
            # Do not silently run code/macros or alter active/closed histories.
            excel.CalculateFull()
            wb.Save()
        except TrialDemoNativeSeedError:
            raise
        except Exception:
            raise TrialDemoNativeSeedError("SAFETY STOP: native Excel seed failed") from None
        finally:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass


def _verify_seed_workbook(
    stage: Path, projection: TrialDemoSourceProjection,
) -> None:
    """Independently re-open and compare every key, time, weekday and provider."""
    with ZipFile(stage) as archive:
        if "xl/vbaProject.bin" not in archive.namelist():
            raise TrialDemoNativeSeedError("VBA project missing after native save")
    wb = load_workbook(stage, read_only=True, data_only=False, keep_vba=True)
    try:
        if not {"PATIENTS", "PATIENT_PLANNER", "PROGRAM_WEEKLY_SLOTS", "DEMO_WEEKLY_VIEW"}.issubset(wb.sheetnames):
            raise TrialDemoNativeSeedError("Native output missing required sheets")
        patients = wb["PATIENTS"]
        headers = {str(value).strip(): col for col, value in enumerate(
            next(patients.iter_rows(min_row=1, max_row=1, values_only=True)), start=1
        ) if value is not None}
        columns = [headers.get(name) for name in projection.patient_registry.headers]
        if any(col is None for col in columns):
            raise TrialDemoNativeSeedError("Native patient projection header missing")
        read_patients = tuple(tuple(patients.cell(row, col).value
                                    for col in columns)
                              for row in range(2, projection.patients + 2))
        if read_patients != projection.patient_registry.rows:
            raise TrialDemoNativeSeedError("Native patient registry differs from projection")
        weekly = wb["PROGRAM_WEEKLY_SLOTS"]
        if tuple(weekly.cell(1, col).value for col in range(1, 7)) != WEEKLY_HEADERS:
            raise TrialDemoNativeSeedError("Native weekly headers differ")
        actual_weekly = []
        for row in range(2, projection.weekday_slot_count + 2):
            values = tuple(weekly.cell(row, col).value for col in range(1, 7))
            actual_weekly.append((
                str(values[0]), values[1], values[2], parse_clock(values[3]),
                values[4], values[5],
            ))
        if tuple(actual_weekly) != projection.weekly_slots.rows:
            raise TrialDemoNativeSeedError("Native weekday slots differ from projection")
        view = wb["DEMO_WEEKLY_VIEW"]
        if view.max_row < projection.weekday_slot_count + 1:
            raise TrialDemoNativeSeedError("Full weekly demo view missing")
        planner = wb["PATIENT_PLANNER"]
        for row in range(2, projection.patients + 2):
            if tuple(planner.cell(row, col).value for col in range(1, 6)) != planner_identity_formulas(row):
                raise TrialDemoNativeSeedError("PATIENT_PLANNER identity formulas differ")
        planner_headers = {
            str(value).strip(): col for col, value in enumerate(
                next(planner.iter_rows(min_row=1, max_row=1, values_only=True)),
                start=1,
            ) if value is not None
        }
        planner_rows = {
            str(row[0]): i for i, row in
            enumerate(projection.patient_registry.rows, start=2)
        }
        for mirror in choose_planner_demo_mirrors(projection.weekly_slots.rows):
            time_key, day_key, provider_key = PLANNER_TREATMENT_HEADERS[mirror.treatment]
            fields = (time_key, day_key) + ((provider_key,) if provider_key else ())
            if any(field not in planner_headers for field in fields):
                raise TrialDemoNativeSeedError("Native planner mirror columns missing")
            row = planner_rows[mirror.patient_id]
            if (
                parse_clock(planner.cell(row, planner_headers[time_key]).value)
                != mirror.start_time
                or planner.cell(row, planner_headers[day_key]).value
                != mirror.day_pattern
                or (
                    provider_key is not None
                    and planner.cell(row, planner_headers[provider_key]).value
                    != mirror.therapist
                )
            ):
                raise TrialDemoNativeSeedError(
                    "Native planner mirror differs from granular source"
                )
    finally:
        wb.close()

    # Independent registry readback through the production PATIENTS parser.
    # This verifies type, MRN, doctor, room, infectious flag and status using
    # the same interpretation that downstream application workflows consume.
    try:
        actual_patients = read_patient_registry(stage)
    except Exception:
        raise TrialDemoNativeSeedError(
            "SAFETY STOP: native PATIENTS registry cannot be read by application"
        ) from None
    expected_patients = tuple(
        (
            str(row[0]), str(row[2]), row[1], row[3] == "Ν", row[4],
            "INPATIENT" if row[5] == "Εσωτερικός" else "OUTPATIENT",
            row[6], row[7],
        )
        for row in projection.patient_registry.rows
    )
    actual_patient_values = tuple(
        (
            patient.patient_id, patient.display_name, patient.room,
            patient.infectious, patient.status,
            patient.patient_type.value.upper(),
            patient.hospital_mrn, patient.responsible_doctor,
        )
        for patient in actual_patients
    )
    if actual_patient_values != expected_patients:
        raise TrialDemoNativeSeedError(
            "SAFETY STOP: application PATIENTS reader differs from approved registry"
        )

    # Independent operational readback through the SAME recurring schedule
    # reader used by daily workload/replacement code, not just sheet equality.
    expected_sessions = Counter(
        (pid, treatment, WEEKDAY_PATTERN[day], clock, provider)
        for pid, treatment, day, clock, provider, _source in
        projection.weekly_slots.rows
    )
    try:
        loaded_sessions = Counter(
            (entry.patient_id, entry.treatment, entry.day_pattern,
             entry.start_time, entry.therapist_id)
            for entry in read_unified_base_schedule(stage)
        )
    except Exception:
        raise TrialDemoNativeSeedError(
            "SAFETY STOP: native weekly entries cannot be read by daily engine"
        ) from None
    if loaded_sessions != expected_sessions:
        raise TrialDemoNativeSeedError(
            "SAFETY STOP: daily engine schedule differs from admitted source"
        )


def seed_trial_demo_workbook(
    plan: TrialImportPlan, *, programme: str | Path,
    review: str | Path, intake: str | Path, output: str | Path,
    clinical_approvals: Mapping[tuple[str, int], ConfirmedClinicalFields],
    backend: TrialDemoSeedBackend | None = None,
) -> TrialDemoNativeSeedReport:
    """Seed one uniquely named, private, fresh DEMO from an EMPTY eligible shell.

    Does not overwrite a file, install new VBA controls, claim interactive
    acceptance or release. Never publish a half-written temp workbook.
    """
    p, r, source, destination = (
        Path(x).resolve() for x in (programme, review, intake, output)
    )
    if (
        not isinstance(plan, TrialImportPlan)
        or not plan.empty_demo_shell or plan.programme_sha256 is None
        or destination.suffix.casefold() != ".xlsm"
        or "_DEMO_" not in destination.name.upper()
        or destination in (p, r, source)
        or destination.exists()
        or not _private_output_target(destination)
    ):
        raise TrialDemoNativeSeedError("SAFETY STOP: unsafe demo output or source plan")
    if backend is None and sys.platform != "win32":
        raise TrialDemoNativeSeedError("Windows Excel required for native seed")

    assert_clean_native_demo_shell(source)
    projection = project_trial_demo_sources(
        plan, programme=p, review=r, intake=source,
        clinical_approvals=clinical_approvals,
    )
    plan.verify_sources(r, source, programme=p)
    # Never create an output inside a tracked repository or a shared production file.
    destination.parent.mkdir(parents=True, exist_ok=True)
    before = _digest(source)
    with tempfile.TemporaryDirectory(prefix="diomil_private_demo_", dir=destination.parent) as folder:
        stage = Path(folder) / "NATIVE_DEMO_STAGE.xlsm"
        with source.open("rb") as original, stage.open("xb") as copied:
            shutil.copyfileobj(original, copied)
        if _digest(stage) != before:
            raise TrialDemoNativeSeedError("SAFETY STOP: fresh demo copy differs")
        if backend is None:
            assert_application_contract(stage)
        (backend or Win32ComTrialDemoSeedBackend()).write(stage, projection)
        _verify_seed_workbook(stage, projection)
        if backend is None:
            # Existing proven native renderer, used only on the private stage.
            refresh_master_projection_in_place(stage)
            apply_master_visual_style_in_place(stage)
            _verify_seed_workbook(stage, projection)
            assert_application_contract(stage)
        plan.verify_sources(r, source, programme=p)
        if _digest(source) != before or destination.exists():
            raise TrialDemoNativeSeedError("SAFETY STOP: source changed or output occupied")
        published_by_this_call = False
        try:
            with destination.open("xb") as published:
                published_by_this_call = True
                with stage.open("rb") as original:
                    shutil.copyfileobj(original, published)
            if _digest(stage) != _digest(destination):
                raise TrialDemoNativeSeedError("SAFETY STOP: published copy mismatch")
            plan.verify_sources(r, source, programme=p)
            if _digest(source) != before:
                raise TrialDemoNativeSeedError("SAFETY STOP: intake shell changed")
        except Exception:
            if published_by_this_call:
                destination.unlink(missing_ok=True)
            raise
    return TrialDemoNativeSeedReport(
        patient_count=projection.patients,
        weekday_slot_count=projection.weekday_slot_count,
        source_unchanged=True,
        # Injected mock backends may validate structural readback only,
        # never claim native Excel/COM acceptance.
        verified_native_readback=backend is None,
        output_path=str(destination),
    )
