from __future__ import annotations

import argparse
from pathlib import Path
import sys

from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parents[2]


def _text(value: object) -> str:
    return str(value).strip() if value is not None else ""


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Inspect current MASTER_SCHEDULE patient ordering without modifying the workbook."
    )
    parser.add_argument("--workbook", type=Path, required=True)
    args = parser.parse_args()

    path = args.workbook.resolve()
    if not path.exists():
        print(f"SAFETY STOP: workbook not found: {path}")
        return 2

    wb = load_workbook(
        path,
        read_only=True,
        data_only=True,
        keep_vba=path.suffix.casefold() == ".xlsm",
    )
    try:
        if "MASTER_SCHEDULE" not in wb.sheetnames:
            print("SAFETY STOP: workbook has no MASTER_SCHEDULE sheet")
            return 2
        if "PATIENTS" not in wb.sheetnames:
            print("SAFETY STOP: workbook has no PATIENTS sheet")
            return 2

        master = wb["MASTER_SCHEDULE"]
        patients = wb["PATIENTS"]

        print("MASTER_SCHEDULE CURRENT ORDER")
        print("row | infectious | room | patient | status")
        master_count = 0
        for row in range(2, master.max_row + 1):
            infectious = _text(master.cell(row, 1).value)
            room = _text(master.cell(row, 2).value)
            patient = _text(master.cell(row, 3).value)
            status = _text(master.cell(row, 11).value)
            if not any((infectious, room, patient, status)):
                continue
            master_count += 1
            print(f"{row} | {infectious} | {room} | {patient} | {status}")

        print()
        print("PATIENTS INPATIENT-LIKE ROWS")
        print("row | id | room | patient | infectious | status | patient_type")
        patient_count = 0

        headers = {
            _text(cell.value): index
            for index, cell in enumerate(patients[1], start=1)
            if _text(cell.value)
        }
        type_col = (
            headers.get("PatientType")
            or headers.get("ΤύποςΑσθενή")
            or headers.get("Τύπος Ασθενή")
        )

        for row in range(2, patients.max_row + 1):
            pid = _text(patients.cell(row, 1).value)
            room = _text(patients.cell(row, 2).value)
            patient = _text(patients.cell(row, 3).value)
            infectious = _text(patients.cell(row, 4).value)
            status = _text(patients.cell(row, 5).value)
            patient_type = _text(patients.cell(row, type_col).value) if type_col else ""
            if not pid and not patient:
                continue
            patient_count += 1
            print(
                f"{row} | {pid} | {room} | {patient} | "
                f"{infectious} | {status} | {patient_type}"
            )

        print()
        print(f"MASTER rows found: {master_count}")
        print(f"PATIENTS rows found: {patient_count}")
        print("READ ONLY: workbook was not modified.")
        return 0
    finally:
        wb.close()


if __name__ == "__main__":
    raise SystemExit(main())
