# Native Excel smoke test

This is the first controlled real write-back test.

It never writes to the v27.1 source workbook. It creates a new preview copy under `Excel/previews/` and writes one visible marker to `THERAPIST_DAILY!B8`.

The tool also verifies:

- the source workbook hash remains unchanged,
- the output can be read back by Microsoft Excel,
- `xl/vbaProject.bin` is still present in the output workbook.

Run from the repository root:

```powershell
python Python/tools/smoke_native_writeback.py
```

Before running, close any open copy of the baseline workbook in Excel.

On success, open only:

`Excel/previews/Rehab_Center_System_v27_1_NATIVE_SMOKE.xlsm`

and confirm that the workbook opens normally and `THERAPIST_DAILY!B8` contains `PYTHON SMOKE TEST`.
