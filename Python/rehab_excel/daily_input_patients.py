from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
import warnings

from openpyxl import load_workbook


@dataclass(frozen=True)
class DailyInputPatientAudit:
    patient_rows: int
    planner_rows: int
    dropdown_names: tuple[str, ...]
    patients_without_planner: tuple[str, ...]
    planner_without_patient: tuple[str, ...]
    identity_mismatches: tuple[str, ...]

    @property
    def ok(self) -> bool:
        return not self.planner_without_patient and not self.identity_mismatches


def _clean(value: object) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _patient_id(value: object) -> str | None:
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    return text or None


def _header_map(ws) -> dict[str, int]:
    first = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    return {
        str(value).strip(): index
        for index, value in enumerate(first)
        if value is not None
    }


def _identity_map(ws, *, sheet_name: str) -> dict[str, str]:
    headers = _header_map(ws)
    required = {"PatientID", "Ασθενής"}
    if not required.issubset(headers):
        missing = ", ".join(sorted(required - set(headers)))
        raise ValueError(f"{sheet_name} missing required columns: {missing}")

    result: dict[str, str] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        pid = _patient_id(row[headers["PatientID"]])
        name = _clean(row[headers["Ασθενής"]])
        if pid is None and name is None:
            continue
        if pid is None or name is None:
            continue
        previous = result.get(pid)
        if previous is not None and previous.casefold() != name.casefold():
            raise ValueError(
                f"{sheet_name} contains conflicting names for PatientID {pid}: "
                f"{previous!r} vs {name!r}"
            )
        result[pid] = name
    return result


def audit_daily_input_patients(path: str | Path) -> DailyInputPatientAudit:
    """Build the DAILY_INPUT patient dropdown from stable patient identities.

    PATIENTS remains the patient authority. PATIENT_PLANNER is used only to
    cross-check PatientID/name identity and to exclude no one silently: patient
    records without a planner row are reported, while planner identities that
    do not exist in PATIENTS are a safety error.
    """

    workbook_path = Path(path)
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message="Data Validation extension is not supported and will be removed",
        )
        wb = load_workbook(
            workbook_path,
            read_only=True,
            data_only=True,
            keep_vba=workbook_path.suffix.casefold() == ".xlsm",
        )
    try:
        for required in ("PATIENTS", "PATIENT_PLANNER"):
            if required not in wb.sheetnames:
                raise ValueError(f"Workbook does not contain {required}")
        patients = _identity_map(wb["PATIENTS"], sheet_name="PATIENTS")
        planner = _identity_map(wb["PATIENT_PLANNER"], sheet_name="PATIENT_PLANNER")
    finally:
        wb.close()

    mismatches: list[str] = []
    for pid in sorted(set(patients) & set(planner), key=lambda x: (len(x), x)):
        if patients[pid].casefold() != planner[pid].casefold():
            mismatches.append(
                f"PatientID {pid}: PATIENTS={patients[pid]!r}, "
                f"PATIENT_PLANNER={planner[pid]!r}"
            )

    patients_without_planner = tuple(
        patients[pid] for pid in patients if pid not in planner
    )
    planner_without_patient = tuple(
        planner[pid] for pid in planner if pid not in patients
    )

    # PATIENTS is the authority. Keep its worksheet order for a predictable
    # dropdown. Cross-checking above prevents a stale planner identity from
    # being mistaken for another patient.
    dropdown_names = tuple(patients.values())

    return DailyInputPatientAudit(
        patient_rows=len(patients),
        planner_rows=len(planner),
        dropdown_names=dropdown_names,
        patients_without_planner=patients_without_planner,
        planner_without_patient=planner_without_patient,
        identity_mismatches=tuple(mismatches),
    )
