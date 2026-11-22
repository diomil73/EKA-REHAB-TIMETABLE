"""Privacy-safe, read-only source audit of infectious flags for MASTER review.

Reports aggregates only, never patient identifiers or flag-by-patient data.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT / "Python") not in sys.path:
    sys.path.insert(0, str(ROOT / "Python"))

from rehab_core.models import PatientType
from rehab_excel.patient_registry_source import read_patient_registry


class InfectiousCoverageError(RuntimeError):
    pass


def count_master_infectious(source: Path, *, reader=read_patient_registry) -> dict[str, int]:
    source = Path(source).resolve()
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise InfectiousCoverageError("Existing demo XLSM source required")
    try:
        patients = tuple(reader(source))
    except Exception as exc:
        raise InfectiousCoverageError("Unable to read demo registry safely") from exc
    inpatients = [p for p in patients if p.patient_type == PatientType.INPATIENT]
    if not inpatients:
        raise InfectiousCoverageError("No inpatient records to verify")
    infectious = sum(bool(p.infectious) for p in inpatients)
    return {
        "total_inpatients": len(inpatients),
        "infectious_inpatients": infectious,
        "noninfectious_inpatients": len(inpatients) - infectious,
    }


def main(argv=None):
    parser = argparse.ArgumentParser(description="Read-only aggregate infectious-flag audit")
    parser.add_argument("--source", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        result = count_master_infectious(args.source)
    except InfectiousCoverageError as exc:
        print("SAFETY STOP: Infectious-flag source audit failed")
        print("DIAGNOSTIC: " + str(exc))
        return 2
    print("MASTER INFECTIOUS SOURCE COVERAGE")
    for key, count in result.items():
        print(f"{key}={count}")
    print("No patient identifiers or clinical row details emitted")
    print("Source opened for read-only inspection; no Excel save")
    print("This does not certify the coloring of an Excel output")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
