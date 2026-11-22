"""Read-only coverage audit for responsible doctors in demo patient registry.

Never emits patient identifiers or names, doctor names, or workbook contents.
No Excel COM, no save, no workbook mutation. Aggregates only.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
for item in (ROOT, ROOT / "Python"):
    if str(item) not in sys.path:
        sys.path.insert(0, str(item))

from rehab_core.models import PatientType
from rehab_excel.patient_registry_source import read_patient_registry


class DoctorCoverageError(RuntimeError):
    pass


def responsible_doctor_coverage(source_path: Path, *, reader=read_patient_registry) -> dict[str, int]:
    path = Path(source_path).resolve()
    if not path.is_file() or path.suffix.casefold() != ".xlsm":
        raise DoctorCoverageError("Existing demo XLSM required")
    try:
        patients = tuple(reader(path))
    except Exception as exc:
        raise DoctorCoverageError("Could not safely read PATIENTS registry") from exc
    inpatients = [p for p in patients if p.patient_type == PatientType.INPATIENT]
    if not inpatients:
        raise DoctorCoverageError("No inpatient records to audit")
    with_doctor = sum(bool(str(p.responsible_doctor or "").strip()) for p in inpatients)
    return {
        "total_inpatients": len(inpatients),
        "inpatients_with_doctor": with_doctor,
        "inpatients_without_doctor": len(inpatients) - with_doctor,
    }


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Privacy-safe, read-only demo responsible-doctor coverage")
    parser.add_argument("--source", required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        result = responsible_doctor_coverage(args.source)
    except DoctorCoverageError as exc:
        print("SAFETY STOP: Doctor coverage audit unavailable")
        print("DIAGNOSTIC: " + str(exc))
        return 2
    print("MASTER RESPONSIBLE DOCTOR SOURCE COVERAGE")
    for name, count in result.items():
        print(f"{name}={count}")
    print("No patient/doctor identifiers emitted")
    print("Source opened read-only; no Excel save or workbook modification")
    print("Coverage is not proof of native MASTER rendering")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
