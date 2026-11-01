from __future__ import annotations

from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[2]
TARGET = REPO_ROOT / "Python" / "rehab_excel" / "daily_input_action_vba.py"


def _replace_once(text: str, old: str, new: str, label: str) -> str:
    count = text.count(old)
    if count != 1:
        raise SystemExit(f"SAFETY STOP: {label} marker count={count}, expected=1")
    return text.replace(old, new, 1)


def main() -> int:
    text = TARGET.read_text(encoding="utf-8")

    old_apply = '    If Not ApplyValidatedDailyInputPayload(payloadPath) Then GoTo CleanUp\n'
    new_apply = (
        '    If Not ApplyValidatedDailyInputPayload(payloadPath) Then\n'
        '        PreserveDailyInputFailureArtifacts payloadPath, responsePath\n'
        '        GoTo CleanUp\n'
        '    End If\n'
    )
    text = _replace_once(text, old_apply, new_apply, "apply failure capture")

    insert_marker = '\nPrivate Function RefreshOperationalViews() As Boolean\n'
    helper = r'''
Private Sub PreserveDailyInputFailureArtifacts(ByVal payloadPath As String, ByVal responsePath As String)
    Dim payloadCopy As String, responseCopy As String

    On Error Resume Next
    payloadCopy = Environ$("TEMP") & "\eka_daily_input_last_payload.tsv"
    responseCopy = Environ$("TEMP") & "\eka_daily_input_last_response.json"

    If Dir$(payloadCopy) <> "" Then Kill payloadCopy
    If Len(payloadPath) > 0 And Dir$(payloadPath) <> "" Then FileCopy payloadPath, payloadCopy

    If Dir$(responseCopy) <> "" Then Kill responseCopy
    If Len(responsePath) > 0 And Dir$(responsePath) <> "" Then FileCopy responsePath, responseCopy
    On Error GoTo 0
End Sub
'''
    if insert_marker not in text:
        raise SystemExit("SAFETY STOP: runtime capture insertion marker not found")
    text = text.replace(insert_marker, "\n" + helper + insert_marker, 1)

    TARGET.write_text(text, encoding="utf-8", newline="\n")
    print(f"UPDATED: {TARGET}")
    print("INSTALLED: DAILY_INPUT failure payload/response capture")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
