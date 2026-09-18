# Native Excel smoke test v2

This smoke test no longer assumes that `THERAPIST_DAILY!B8` is empty.

It opens the baseline workbook read-only, searches only the confirmed daily grids
`B2:J8` and `B12:J18`, and selects an empty cell that is not a formula or merged
cell. It then copies the `.xlsm`, writes `PYTHON SMOKE TEST` only to the copy,
reopens the copy to verify the value, confirms that `xl/vbaProject.bin` still
exists, and confirms that the source SHA-256 did not change.

Run with Excel closed:

```powershell
python Python/tools/smoke_native_writeback.py
```

If an older preview already exists:

```powershell
python Python/tools/smoke_native_writeback.py --overwrite
```

A successful run prints the exact automatically selected cell. Open only the
preview workbook under `Excel/previews/` for manual inspection.
