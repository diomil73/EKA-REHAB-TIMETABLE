from __future__ import annotations

from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[2]
TARGET = REPO_ROOT / "Python" / "rehab_excel" / "operator_visibility_patch_vba.py"
DIAGNOSTIC = REPO_ROOT / "Python" / "tools" / "inspect_native_daily_restore_contract.py"

OLD_VERIFY = r'''Private Function VerifyCharacterStrike(ByVal cell As Range, ByVal enabled As Boolean) As Boolean
    Dim textLength As Long, charIndex As Long
    Dim strikeValue As Variant

    textLength = Len(CStr(cell.Value))
    If textLength = 0 Then
        strikeValue = cell.Font.Strikethrough
        If IsNull(strikeValue) Then Exit Function
        VerifyCharacterStrike = (CBool(strikeValue) = enabled)
        Exit Function
    End If

    For charIndex = 1 To textLength
        strikeValue = cell.Characters(charIndex, 1).Font.Strikethrough
        If IsNull(strikeValue) Then Exit Function
        If CBool(strikeValue) <> enabled Then Exit Function
    Next charIndex
    VerifyCharacterStrike = True
End Function'''

NEW_VERIFY = r'''Private Function VerifyCharacterStrike(ByVal cell As Range, ByVal enabled As Boolean) As Boolean
    Dim textLength As Long, charIndex As Long
    Dim strikeValue As Variant

    ' Prefer Excel's whole-cell answer whenever it is definitive.  Some MASTER
    ' cells expose Font.Strikethrough correctly but do not expose Characters()
    ' through COM; forcing character read-back there creates a false mismatch.
    strikeValue = cell.Font.Strikethrough
    If Not IsNull(strikeValue) Then
        VerifyCharacterStrike = (CBool(strikeValue) = enabled)
        Exit Function
    End If

    ' Only mixed-format cells that return Null need character-level fallback.
    textLength = Len(CStr(cell.Value))
    If textLength = 0 Then Exit Function

    On Error GoTo CharacterReadFailed
    For charIndex = 1 To textLength
        strikeValue = cell.Characters(charIndex, 1).Font.Strikethrough
        If IsNull(strikeValue) Then Exit Function
        If CBool(strikeValue) <> enabled Then Exit Function
    Next charIndex
    VerifyCharacterStrike = True
    Exit Function

CharacterReadFailed:
    VerifyCharacterStrike = False
End Function'''

OLD_DIAGNOSTIC = '"daily.status_character_clear": ".Characters(1, Len(CStr(master.Cells(row, statusCol).Value))).Font.Strikethrough = False" in daily,'
NEW_DIAGNOSTIC = '"daily.status_character_clear": "ApplyCharacterStrike master.Cells(row, statusCol), False" in daily,'


def _replace_exact(path: Path, old: str, new: str, label: str) -> None:
    text = path.read_text(encoding="utf-8")
    count = text.count(old)
    if count != 1:
        raise SystemExit(f"SAFETY STOP: {label} marker count={count}, expected=1")
    path.write_text(text.replace(old, new, 1), encoding="utf-8")


def main() -> int:
    _replace_exact(TARGET, OLD_VERIFY, NEW_VERIFY, "VerifyCharacterStrike")
    _replace_exact(DIAGNOSTIC, OLD_DIAGNOSTIC, NEW_DIAGNOSTIC, "status diagnostic")
    print("MASTER STRIKE READ-BACK SOURCE FIX APPLIED")
    print(f"Updated: {TARGET.relative_to(REPO_ROOT)}")
    print(f"Updated: {DIAGNOSTIC.relative_to(REPO_ROOT)}")
    print("No workbook was opened or modified.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
