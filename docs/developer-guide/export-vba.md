# Export VBA from an Excel add-in

`tools/export_vba.py` exports the complete VBA project from a NADABAS `.xlam`,
`.xlsm`, or `.xlsb` file into the repository layout used below `vba/`.

## Prerequisites

- Windows with Microsoft Excel installed
- Python 3.12 or newer and the project dependencies installed
- An unlocked VBA project in the source file
- In Excel: **File > Options > Trust Center > Trust Center Settings > Macro
  Settings > Trust access to the VBA project object model**

Close other Excel windows before exporting.

## Export

From the repository root, export into a new staging directory:

```powershell
python tools/export_vba.py C:\path\to\NADABAS.xlam `
  --output-root build\exported-vba
```

The source workbook is opened read-only with macros, events, alerts, and link
updates disabled. Components are first written to a temporary directory; the
requested output becomes visible only after the complete export succeeds.

The output contains:

- standard modules in `modules/*.bas`;
- class and document modules in `classes/*.cls`; and
- complete UserForms in `forms/*.frm` with their binary `.frx` sidecars.

Use `--force` only when an existing staging directory may be replaced. Copy
the reviewed `modules/`, `classes/`, and `forms/` directories into `vba/`, but
keep `vba/ReadMe.md`.

## Validate before publishing

Run the unit tests and security scan, then verify a rebuild through Excel:

```powershell
python -m unittest discover -s tests -v
python tools/scan_vba_security.py vba
python tools/import_vba.py C:\path\to\NADABAS.xlam `
  --source-root vba `
  --dry-run `
  --replace-form-designers
```

Exporting does not modify or invalidate the signature on the source workbook.
A rebuilt workbook is a new artifact and must be tested and signed separately.
