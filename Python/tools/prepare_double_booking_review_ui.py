from __future__ import annotations

import argparse
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from tools import build_double_booking_review_payload  # noqa: E402
from tools.export_double_booking_review_rows import export_review_rows  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build double-booking review JSON and UI TSV in one Python process."
    )
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--review", required=True, type=Path)
    parser.add_argument("--rows", required=True, type=Path)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    input_book = args.input.resolve()
    review = args.review.resolve()
    rows = args.rows.resolve()
    if not input_book.exists():
        print(f"SAFETY STOP: input workbook not found: {input_book}")
        return 2
    if not args.overwrite and (review.exists() or rows.exists()):
        print("SAFETY STOP: review or rows output already exists")
        return 2

    old_argv = sys.argv[:]
    try:
        sys.argv = [
            "build_double_booking_review_payload.py",
            "--input",
            str(input_book),
            "--output",
            str(review),
            "--overwrite",
        ]
        rc = build_double_booking_review_payload.main()
    finally:
        sys.argv = old_argv

    if rc != 0:
        return rc

    try:
        count = export_review_rows(review, rows)
    except Exception as exc:
        try:
            rows.unlink(missing_ok=True)
        except OSError:
            pass
        print(f"SAFETY STOP: {exc}")
        return 2

    print("DOUBLE BOOKING REVIEW UI PREP OK")
    print(f"Options exported: {count}")
    print(f"Review: {review}")
    print(f"Rows: {rows}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
