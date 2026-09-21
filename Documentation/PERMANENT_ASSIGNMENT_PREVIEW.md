# Permanent Assignment Preview

Creates a **new `.xlsm` preview copy** and changes one physiotherapy assignment
in `PATIENT_PLANNER` after the weekly capacity/conflict gate passes.

The source workbook is never edited. The preview preserves the VBA project.

Current development workflow:
1. validate weekly recurring capacity/conflicts,
2. write the approved test move only to a preview copy,
3. open that preview in Excel,
4. use the existing Refresh/ΑΝΑΝΕΩΣΗ action to rebuild operational views,
5. inspect the resulting timetable before any future production write-back is enabled.

This is a development preview tool, not yet the final editor workflow.
