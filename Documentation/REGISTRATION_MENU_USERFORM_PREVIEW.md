# Registration menu UserForm preview

This stage adds the first visible central registration UI shell to a safe `.xlsm` preview copy.

## What it creates

The installer adds two VBA components to the preview workbook:

- `modRegistrationMenu`
- `frmRegistrationMenu`

The public macro is:

- `ShowRegistrationMenu`

Running that macro opens a centered UserForm titled `Κεντρικό Μενού Εγγραφών` with four buttons:

1. `Νέος ασθενής`
2. `Νέος θεραπευτής`
3. `Νέος φοιτητής`
4. `Κλείσιμο`

At this stage the three registration buttons intentionally show informational placeholder messages only. They do not yet write registration data. The next stage will connect them to the already-proven registration menu contract and orchestration layer.

## Safety model

`create_registration_menu_preview()`:

- requires an existing `.xlsm` source;
- requires a different `.xlsm` output path;
- hashes the source before any work;
- creates a new copy;
- opens only the copy in Microsoft Excel;
- installs the VBA module/UserForm into the copy;
- saves and verifies the expected VBA components;
- hashes the source again and fails if it changed;
- removes a failed preview when possible.

The baseline workbook is never opened for writing by this workflow.

## Windows / Excel requirement

Programmatic creation of a VBA UserForm requires:

- Windows;
- Microsoft Excel desktop;
- `pywin32`;
- Excel permission to edit the VBA project programmatically.

If Excel blocks VBProject access, enable:

`File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model`

This setting is needed only for the development installer that creates the preview UserForm. It is not a replacement for normal macro security.

## Build the preview

From the repository root:

```powershell
python Python/tools/build_registration_menu_preview.py --overwrite
```

Default output:

`Excel/previews/REGISTRATION_MENU_PREVIEW.xlsm`

Expected success output includes:

- `Source unchanged: True`
- `VBA project present: True`
- `Menu module present: True`
- `Menu form present: True`

Then open the preview workbook in Excel, enable macros for the trusted development copy, press `Alt+F8`, select `ShowRegistrationMenu`, and run it.

## Current boundary

This PR proves the visible central menu can be safely installed in a copied workbook while preserving the source. It does not yet connect the buttons to patient, therapist, or student registration execution.
