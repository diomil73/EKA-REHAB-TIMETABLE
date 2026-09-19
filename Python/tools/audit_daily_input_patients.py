from __future__ import annotations

import argparse
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.daily_input_patients import audit_daily_input_patients  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit DAILY_INPUT patient identities.")
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    args = parser.parse_args()
    audit = audit_daily_input_patients(args.source.resolve())
    print("DAILY INPUT PATIENT AUDIT")
    print(f"PATIENTS identities: {audit.patient_rows}")
    print(f"PATIENT_PLANNER identities: {audit.planner_rows}")
    print(f"Dropdown names: {len(audit.dropdown_names)}")
    print(f"Patients without planner identity: {len(audit.patients_without_planner)}")
    for name in audit.patients_without_planner:
        print(f"  PATIENTS only: {name}")
    print(f"Planner identities missing from PATIENTS: {len(audit.planner_without_patient)}")
    for name in audit.planner_without_patient:
        print(f"  PLANNER only: {name}")
    print(f"Identity mismatches: {len(audit.identity_mismatches)}")
    for item in audit.identity_mismatches:
        print(f"  {item}")
    print(f"Audit OK: {audit.ok}")
    return 0 if audit.ok else 2


if __name__ == "__main__":
    raise SystemExit(main())
