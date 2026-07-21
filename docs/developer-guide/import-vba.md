# Import VBA into an XLAM

`tools/import_vba.py` builds a new NADABAS add-in from an existing `.xlam`
template and the version-controlled files below `vba/`. The template is never
modified.

## Prerequisites

- Windows with Microsoft Excel installed
- Python 3.12 or newer and the project dependencies installed
- An unlocked VBA project in the template
- In Excel: **File > Options > Trust Center > Trust Center Settings > Macro
  Settings > Trust access to the VBA project object model**

Close the NADABAS add-in before building if Excel has locked the file.

## Validate first

From the repository root, run:

```powershell
python tools/import_vba.py C:\path\to\NADABAS.xlam --source-root vba --dry-run
```

The dry run opens the template with macros and events disabled, validates the
sources, and lists every planned component operation without changing a file.

## Build the add-in

```powershell
python tools/import_vba.py C:\path\to\NADABAS.xlam `
  --source-root vba `
  --output build\NADABAS-review.xlam
```

Use `--force` only when an existing output file may be replaced. The importer
works on a temporary copy and moves it to the output path only after Excel has
saved successfully.

## Modules, classes, and UserForms

For a component already present in the template, the importer replaces only
its code module. This is intentional: it preserves the controls and binary
designer state of existing UserForms. It also lets the code-only `.frm` files
currently stored in this repository be applied safely.

A new standard module can be created from a `.bas` file. A new class or
UserForm must be a complete export when it relies on VBA attributes or form
designer data. A new UserForm therefore needs both its complete `.frm` file and
every referenced `.frx` sidecar.

To apply reviewed UserForm designer changes as well as code, export a complete
`.frm`/`.frx` pair and build with:

```powershell
python tools/import_vba.py C:\path\to\NADABAS.xlam `
  --source-root vba `
  --output build\NADABAS-review.xlam `
  --replace-form-designers
```

## Suggested GitHub workflow

1. Export or update VBA source files on a feature branch.
2. Run the security scan and unit tests.
3. Open a pull request and review the text-based VBA diff.
4. Merge the approved pull request.
5. Check out the approved commit and run the importer against the official
   NADABAS template.
6. Open the generated add-in in Excel and run the functional test checklist.
7. Apply the Statistics Norway digital signature after testing, then publish
   the signed artifact through the normal release channel.

Editing VBA invalidates an existing digital signature. The generated file must
therefore be signed after the import and verification steps.

## Troubleshooting

- **Excel denied access to the VBA project:** enable **Trust access to the VBA
  project object model**, restart Excel, and try again.
- **The VBA project is locked:** unlock the project in the template before
  importing.
- **Missing UserForm binary sidecar:** export and commit the `.frx` file next to
  its `.frm` file.
- **Code-only UserForm is not present in the template:** use a template that
  already contains the form, or create a complete `.frm`/`.frx` export.
- **Output exists:** choose a new output path or pass `--force` deliberately.
