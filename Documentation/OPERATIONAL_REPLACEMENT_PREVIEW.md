# Operational replacement preview

This is a deliberately narrow native-Excel preview. It does **not** alter the
baseline workbook.

Scenario used for the first visual check:

- patient: `ΖΑΛΟΚΩΣΤΑΣ`
- original provider: `Αργέντος`
- original time: `12:15`
- replacement provider: `Φιλιππούσης`
- replacement time: `12:15`

The tool resolves the real `THERAPIST_DAILY` cells from the workbook layout,
creates a new `.xlsm` copy, strikes through the original patient line, and adds
the patient to the replacement provider cell. The infectious presentation is
kept visible through the yellow semantic formatting.

Run from the repository root with Excel closed:

```powershell
python Python/tools/preview_operational_replacement.py --overwrite
```

The preview is written under `Excel/previews/`, which must remain ignored by
Git.
