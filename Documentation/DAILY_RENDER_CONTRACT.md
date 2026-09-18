# Daily Excel Render Contract

This layer converts `DailySessionState` into a **read-only semantic render plan** for the existing workbook layout.

## Safety rules

- `PATIENT_PLANNER` remains the base schedule source and is never changed by this layer.
- `MASTER_SCHEDULE` is located for traceability, but this stage does **not** plan value changes there because it contains recurring day-pattern text.
- Daily visual changes are planned only for `THERAPIST_DAILY`.
- The planner never opens or saves the workbook in write mode.
- Missing or ambiguous patient/provider mappings are reported as errors rather than guessed.

## Daily rendering semantics

- Active session: normal patient line.
- Patient absent: original patient line remains in its original slot with strike-through.
- Therapist absent: original patient line remains in its original slot with strike-through; no replacement slot is invented.
- Accepted replacement: original line is struck through and the patient is also planned in the effective provider/time cell.
- Multiple patients in the same legacy slot are preserved as separate lines in the same cell.

## Visual roles

- Robotic: pink fill.
- Infectious: yellow indicator. If a session is both robotic and infectious, pink remains the fill and the infectious signal is preserved as a yellow border.
- Minimum planned font size: 10.
- Wrap text remains enabled.

## Students

A student replacement can resolve the current legacy header (`φοιτ N`) while the student model keeps the real name and placement dates. Header renaming/green-font application is deliberately left for the next controlled UI step.
