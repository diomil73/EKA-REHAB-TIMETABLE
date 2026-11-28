# DIOMIL — PowerShell Operator Execution Standard

**Status:** Mandatory operator procedure  
**Applies to:** All Windows PowerShell instructions for DIOMIL  
**Last updated:** 2026-11-28

## Why this exists

The project has encountered the same failure more than once: a multi-line PowerShell compound block was pasted/executed in pieces. PowerShell returned to a new prompt before the intended block was complete, so a later else or finally was interpreted as a standalone command. A guarded workflow then continued incorrectly or its report/log handling failed. On 2026-11-28, an active Start-Transcript also prevented moving/renaming its own open file.

This is an execution-wrapper failure, not proof that the underlying Python test or workbook candidate passed or failed.

## Mandatory rules

1. **Long or safety-critical workflow: use a saved script.** Provide one complete uniquely named .ps1 file and run it once:
   ~~~powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\DIOMIL_<purpose>_<timestamp>.ps1
   ~~~
   Use process-scoped bypass only when permitted/needed by local IT policy. Do not globally weaken execution policy.

2. **Short workflow: one complete scriptblock.** The operator pastes the entire & { ... } from its opening to closing brace in one submission. Never split it across separate terminal submissions.

3. Never send or ask the operator to run an isolated else, catch, finally, closing brace, or here-string terminator. Do not provide independent follow-on commands that could execute after a prior gate failed.

4. A multi-line try/catch/finally is permitted only inside a complete saved .ps1 or a complete one-submission & { ... }. If a new PowerShell prompt appears before the expected closing brace/result, classify the run as INCOMPLETE / SAFETY STOP. Do not continue from the middle.

5. Capture a native executable's exit code immediately:
   ~~~powershell
   python <tool-and-args>
   $toolExit = $LASTEXITCODE
   if ($toolExit -ne 0) { throw "SAFETY STOP: tool exit $toolExit" }
   ~~~
   Do not run another native executable before capturing the prior code. Keep PowerShell errors and native-process exit codes distinct.

6. **Transcript lifecycle:** if Start-Transcript is used, call Stop-Transcript successfully before reading, moving, copying, or renaming that transcript. Never move an active transcript. For simple output capture, prefer redirection to a uniquely named log instead of a transcript.

7. A report status is written only after all gates and exit codes have been checked. If the wrapper itself breaks, the result is S (SAFETY STOP), not a test B. A B is valid only when the diagnostic completed and explicitly reported a negative finding. A missing report or stale variable cannot be interpreted as success or failure evidence.

8. Preserve logs and failed candidates. Never delete, overwrite, or retry a frozen artifact unless the project guide explicitly authorizes a new uniquely named candidate.

## Correct patterns

### Preferred: saved .ps1

Save the entire script, including all braces, here-strings, error handling and final status handling, then invoke it once with -File. The script owns its full control flow. It should stop at the first failed gate and leave the first diagnostic intact.

### Acceptable: one complete interactive block

Paste the complete block in one operation:

~~~powershell
& {
    $ErrorActionPreference = "Stop"
    $status = "S"

    # All validation and commands belong inside this same block.
    # Capture $LASTEXITCODE immediately after each native command.
    # Do not submit any part of this block separately.

    $status = "G"
    "FINAL_STATUS=$status"
}
~~~

This is a structural example only, not a DIOMIL test command.

## Incident references

- Documentation/DIOMIL_MASTER_CONTINUITY_GUIDE.md §3.3 and §4.253.
- Documentation/PROJECT_STATUS.md, 2026-11-08 incident: interactive continuation after a terminating throw, false-positive success banner, and subsequent one-block runbook requirement.
- Documentation/PROJECT_STATUS.md, 2026-11-28 incident: split finally, active transcript file lock and invalid FINAL_STATUS=B.
- Documentation/CURRENT_HANDOFF.md: current operator status and next action.

## Current 2026-11-28 disposition

The H10 K schema comparison of three frozen candidates is PENDING because its PowerShell wrapper did not complete. The reported FINAL_STATUS=B is invalid and must not mark any candidate BAD. Re-run only the read-only schema comparison using this standard. Earlier PREPARED digest test (1 passed) and warning diagnostic (109 passed, 1 skipped) remain green and must not be repeated without a technical reason.
