from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path

EXPECTED_SCHEMA = "eka.double_booking_review.v1"
HEADERS = (
    "target_session_id",
    "patient_name",
    "original_time",
    "therapist_id",
    "therapist_name",
    "replacement_time",
    "warning",
)


def export_review_rows(review_path: Path, output_path: Path) -> int:
    payload = json.loads(review_path.read_text(encoding="utf-8-sig"))
    if payload.get("schema") != EXPECTED_SCHEMA:
        raise ValueError("Unsupported double-booking review schema")
    if payload.get("requires_user_confirmation") is not True:
        raise ValueError("Review payload does not require explicit confirmation")

    rows: list[dict[str, str]] = []
    for item in payload.get("items", []):
        target_session_id = str(item.get("target_session_id", "")).strip()
        patient_name = str(item.get("patient_name", "")).strip()
        original_time = str(item.get("original_time", "")).strip()
        if not target_session_id or not patient_name or not original_time:
            raise ValueError("Review item is missing required UI fields")
        for option in item.get("options", []):
            therapist_id = str(option.get("therapist_id", "")).strip()
            therapist_name = str(option.get("therapist_name", "")).strip()
            replacement_time = str(option.get("replacement_time", "")).strip()
            warning = str(option.get("warning", "")).strip()
            if not therapist_id or not therapist_name or not replacement_time or not warning:
                raise ValueError("Review option is missing required UI fields")
            rows.append(
                {
                    "target_session_id": target_session_id,
                    "patient_name": patient_name,
                    "original_time": original_time,
                    "therapist_id": therapist_id,
                    "therapist_name": therapist_name,
                    "replacement_time": replacement_time,
                    "warning": warning.replace("\t", " ").replace("\r", " ").replace("\n", " "),
                }
            )

    if not rows:
        raise ValueError("Review payload contains no double-booking options")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=HEADERS, delimiter="\t", lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
    return len(rows)


def main() -> int:
    parser = argparse.ArgumentParser(description="Export double-booking review choices as UI-friendly TSV.")
    parser.add_argument("--review", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    review = args.review.resolve()
    output = args.output.resolve()
    if not review.exists():
        print(f"SAFETY STOP: review payload not found: {review}")
        return 2
    if output.exists() and not args.overwrite:
        print(f"SAFETY STOP: output already exists: {output}")
        return 2

    try:
        count = export_review_rows(review, output)
    except (ValueError, json.JSONDecodeError, OSError) as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("DOUBLE BOOKING UI ROWS OK")
    print(f"Options exported: {count}")
    print(f"Output: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
