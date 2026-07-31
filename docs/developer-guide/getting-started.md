# Getting started

NADABAS keeps the exported VBA project, supporting Python tools, tests, and
technical documentation in Git. The exported files below `vba/` are the source
of truth; Excel workbooks and add-ins are build or release artifacts.

## Prerequisites

- Python 3.12 or newer
- Poetry 2.2 or newer
- Git and pre-commit
- Windows and Microsoft Excel for VBA import and export operations

## Repository structure

```text
vba/
  modules/
  classes/
  forms/
tools/
tests/
docs/
```

Install the project and development dependencies:

```powershell
poetry install
poetry run pre-commit install
```

## Exporting VBA

Export a complete VBA project to a staging directory before copying reviewed
changes into `vba/`. See [Export VBA](export-vba.md).

## Importing VBA

Dry-run an import and build a review copy from the version-controlled source.
The template is never modified. See [Import VBA](import-vba.md).

## Validate a change

Run the same core checks used by GitHub Actions:

```powershell
poetry run pre-commit run --all-files
poetry run pytest -v
poetry run python tools/scan_vba_security.py vba `
  --approvals tools/vba-security-approvals.json `
  --show-approved
```

Read the repository's
[contribution guidelines](https://github.com/statisticsnorway/nadabas/blob/creation/CONTRIBUTING.md)
before opening a pull request.
