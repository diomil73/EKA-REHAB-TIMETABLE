"""Private companion checklist; no approval, source edits or import."""
from __future__ import annotations
import argparse
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "Python"))
from rehab_excel.trial_review_checklist import create_private_review_checklist


def main(argv=None):
    parser = argparse.ArgumentParser(description="Diomil private pending-review inventory")
    for option in ("programme", "review", "output"):
        parser.add_argument("--" + option, required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        counts = create_private_review_checklist(
            programme=args.programme, review=args.review, output=args.output,
        )
    except Exception as exc:
        print(f"SAFETY STOP: checklist failed ({type(exc).__name__}).")
        return 2
    print("PRIVATE PENDING REVIEW ONLY; NO APPROVAL OR IMPORT")
    for category, count in sorted(counts.items()):
        print(f"{category}: {count}")
    print("Original sources unchanged; all tasks pending.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
