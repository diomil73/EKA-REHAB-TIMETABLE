from __future__ import annotations

from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[2]
TARGET = REPO_ROOT / "Python" / "rehab_excel" / "operator_visibility_patch_vba.py"


def main() -> int:
    text = TARGET.read_text(encoding="utf-8")

    old_verify = '''    verify_replacement = '''\''    If PayloadValue(parts(5)) = "1" And Len(textRuns) = 0 Then
        If Not VerifyCharacterStrike(cell, True) Then Exit Function
    End If'''\'''
'''
    new_verify = '''    verify_replacement = '''\''    If PayloadValue(parts(5)) = "1" And Len(textRuns) = 0 Then
        If Not VerifyCharacterStrike(cell, True) Then
            If PayloadValue(parts(11)) <> "master_daily_non_active_overlay" Then Exit Function
            strikeValue = cell.Font.Strikethrough
            If Not IsNull(strikeValue) Then Exit Function
        End If
    End If'''\'''
'''
    if old_verify not in text:
        raise SystemExit("SAFETY STOP: generic strike verification marker not found exactly as expected")
    text = text.replace(old_verify, new_verify, 1)

    old_source = '''    source_verify_replacement = '''\''        Case "master_daily_non_active_overlay"
            VerifySourceTagPresentation = _
                (VerifyCharacterStrike(cell, True) And CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))
        Case "master_daily_patient_status_overlay"
            VerifySourceTagPresentation = _
                (VerifyCharacterStrike(cell, False) And CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))'''\'''
'''
    new_source = '''    source_verify_replacement = '''\''        Case "master_daily_non_active_overlay"
            strikeValue = cell.Font.Strikethrough
            If IsNull(strikeValue) Then
                VerifySourceTagPresentation = (CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))
            Else
                VerifySourceTagPresentation = _
                    (CBool(strikeValue) = True And CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))
            End If
        Case "master_daily_patient_status_overlay"
            VerifySourceTagPresentation = _
                (VerifyCharacterStrike(cell, False) And CLng(cell.Font.Color) = CLng(RGB(255, 0, 0)))'''\'''
'''
    if old_source not in text:
        raise SystemExit("SAFETY STOP: source-tag strike verification marker not found exactly as expected")
    text = text.replace(old_source, new_source, 1)

    TARGET.write_text(text, encoding="utf-8", newline="\n")
    print(f"UPDATED: {TARGET}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
