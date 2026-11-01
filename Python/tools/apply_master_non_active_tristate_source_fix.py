from __future__ import annotations

from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[2]
TARGET = REPO_ROOT / "Python" / "rehab_excel" / "operator_visibility_patch_vba.py"


def _replace_once(text: str, old: str, new: str, label: str) -> str:
    count = text.count(old)
    if count != 1:
        raise SystemExit(f"SAFETY STOP: {label} marker count={count}, expected=1")
    return text.replace(old, new, 1)


def main() -> int:
    text = TARGET.read_text(encoding="utf-8")

    old_verify = (
        "    verify_replacement = '''    If PayloadValue(parts(5)) = \"1\" And Len(textRuns) = 0 Then\n"
        "        If Not VerifyCharacterStrike(cell, True) Then Exit Function\n"
        "    End If'''\n"
    )
    new_verify = (
        "    verify_replacement = '''    If PayloadValue(parts(5)) = \"1\" And Len(textRuns) = 0 Then\n"
        "        If Not VerifyCharacterStrike(cell, True) Then\n"
        "            If PayloadValue(parts(11)) <> \"master_daily_non_active_overlay\" Then Exit Function\n"
        "            strikeValue = cell.Font.Strikethrough\n"
        "            If Not IsNull(strikeValue) Then Exit Function\n"
        "        End If\n"
        "    End If'''\n"
    )
    text = _replace_once(text, old_verify, new_verify, "generic strike verification")

    old_source = (
        "    source_verify_replacement = '''        Case \"master_daily_non_active_overlay\"\n"
        "            VerifySourceTagPresentation = _\n"
        "                (VerifyCharacterStrike(cell, True) And CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))\n"
        "        Case \"master_daily_patient_status_overlay\"\n"
        "            VerifySourceTagPresentation = _\n"
        "                (VerifyCharacterStrike(cell, False) And CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))'''\n"
    )
    new_source = (
        "    source_verify_replacement = '''        Case \"master_daily_non_active_overlay\"\n"
        "            strikeValue = cell.Font.Strikethrough\n"
        "            If IsNull(strikeValue) Then\n"
        "                VerifySourceTagPresentation = (CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))\n"
        "            Else\n"
        "                VerifySourceTagPresentation = _\n"
        "                    (CBool(strikeValue) = True And CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))\n"
        "            End If\n"
        "        Case \"master_daily_patient_status_overlay\"\n"
        "            VerifySourceTagPresentation = _\n"
        "                (VerifyCharacterStrike(cell, False) And CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))'''\n"
    )
    text = _replace_once(text, old_source, new_source, "source-tag strike verification")

    TARGET.write_text(text, encoding="utf-8", newline="\n")
    print(f"UPDATED: {TARGET}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
