from __future__ import annotations

from hashlib import sha256
from pathlib import Path
import argparse
import re
import sys


def _sha256(path: Path) -> str:
    h = sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser(description="Inspect installed DAILY_INPUT runtime provenance without modifying workbook")
    parser.add_argument("--workbook", required=True, type=Path)
    args = parser.parse_args()

    workbook_path = args.workbook.resolve()
    if not workbook_path.exists():
        print(f"SAFETY STOP: workbook not found: {workbook_path}")
        return 2
    if sys.platform != "win32":
        print("SAFETY STOP: Windows/Excel required")
        return 2

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        print("SAFETY STOP: pywin32 is required")
        return 2

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(workbook_path), UpdateLinks=0, ReadOnly=True)
        component = workbook.VBProject.VBComponents("modDailyInputAction")
        code_module = component.CodeModule
        code = str(code_module.Lines(1, int(code_module.CountOfLines)))

        print("WORKBOOK:", workbook_path)
        print("WORKBOOK_SHA256:", _sha256(workbook_path))
        print("MODULE_LINES:", int(code_module.CountOfLines))
        print("HAS_CHARACTER_STRIKE_HELPER:", "Private Function VerifyCharacterStrike" in code)
        print("HAS_TRISTATE_NON_ACTIVE_BRANCH:", 'If PayloadValue(parts(11)) <> "master_daily_non_active_overlay" Then Exit Function' in code)
        print("HAS_WHOLE_CELL_FIRST_COMMENT:", "Prefer Excel's whole-cell answer whenever it is definitive" in code)

        resolver_match = re.search(
            r"Private Function ResolveDailyInputPreparationScript\(\) As String(?P<body>.*?)End Function",
            code,
            re.IGNORECASE | re.DOTALL,
        )
        if resolver_match:
            print("RESOLVER_PRESENT: True")
            print("RESOLVER_BODY_BEGIN")
            print(resolver_match.group("body").strip())
            print("RESOLVER_BODY_END")
        else:
            print("RESOLVER_PRESENT: False")

        candidates = [
            workbook_path.parent / "Python" / "tools" / "prepare_daily_input_in_place.py",
            (workbook_path.parent / ".." / ".." / "Python" / "tools" / "prepare_daily_input_in_place.py").resolve(),
        ]
        chosen = None
        for index, candidate in enumerate(candidates, start=1):
            exists = candidate.exists()
            print(f"CANDIDATE_{index}:", candidate)
            print(f"CANDIDATE_{index}_EXISTS:", exists)
            if exists:
                print(f"CANDIDATE_{index}_SHA256:", _sha256(candidate))
                if chosen is None:
                    chosen = candidate
        print("RESOLVED_SCRIPT:", chosen or "<none>")
        if chosen:
            print("RESOLVED_SCRIPT_SHA256:", _sha256(chosen))

        repo_script = (Path(__file__).resolve().parents[2] / "Python" / "tools" / "prepare_daily_input_in_place.py").resolve()
        print("REPO_SCRIPT:", repo_script)
        print("REPO_SCRIPT_SHA256:", _sha256(repo_script))
        print("RESOLVED_EQUALS_REPO:", bool(chosen and chosen.resolve() == repo_script))
        print("RESOLVED_HASH_EQUALS_REPO:", bool(chosen and _sha256(chosen) == _sha256(repo_script)))
        return 0
    except Exception as exc:
        print(f"SAFETY STOP: provenance inspection failed: {exc}")
        return 2
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
