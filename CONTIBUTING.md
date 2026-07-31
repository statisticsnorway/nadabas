# Contributing to NADABAS

Thank you for helping improve NADABAS. This repository contains the
version-controlled VBA source, supporting Python tools, tests, and technical
documentation for an Excel-based production system. Changes must therefore be
reviewable, reproducible, and safe to test without production data.

## Before you start

- Search the existing issues before opening a new one.
- Discuss broad architectural, database, or user-interface changes with the
  maintainers before investing in an implementation.
- Use the private vulnerability reporting link in [SECURITY.md](SECURITY.md)
  for suspected vulnerabilities. Do not disclose them in a public issue.
- Never commit passwords, tokens, certificates, connection strings containing
  credentials, personal data, production data, or signed release artifacts.

## Set up the development environment

NADABAS requires Python 3.12 or newer. The Excel import and export workflows
additionally require Windows, Microsoft Excel, and an unlocked VBA project.

Install the project and development dependencies from the repository root:

```powershell
poetry install
poetry run pre-commit install
```

Run the complete local validation suite before opening a pull request:

```powershell
poetry run pre-commit run --all-files
poetry run pytest -v
poetry run python tools/scan_vba_security.py vba `
  --approvals tools/vba-security-approvals.json `
  --show-approved
```

The GitHub Actions workflows are authoritative if a local result differs from
CI.

## Work on a focused branch

Create a branch from the branch named by the issue or maintainer. Keep each
branch focused on one problem. Use a descriptive name such as
`fix/import-userforms`, `feature/batch-validation`, or
`docs/database-architecture`.

GitHub currently identifies `main` as the default branch, while some active
development is integrated through `creation`. Confirm the intended target
branch before opening a pull request until the branch strategy is consolidated.

## Change VBA safely

The exported files below `vba/` are the source of truth. An Excel add-in or
workbook is a build or release artifact, not a substitute for the exported
source.

1. Export the VBA project to a staging directory as described in
   [Export VBA](docs/developer-guide/export-vba.md).
2. Review the staged diff before copying intended changes into `vba/`.
3. Preserve `.bas`, `.cls`, and `.frm` files with CRLF line endings. Keep each
   required `.frx` file beside its `.frm` file.
4. Run pre-commit, tests, and the VBA security scanner.
5. Dry-run an import into a copy of the official template as described in
   [Import VBA](docs/developer-guide/import-vba.md).
6. Record the manual Excel checks performed in the pull request.

Do not weaken a security rule or add a broad approval to make a scan pass. A
necessary exception must match the exact path, rule, and expected line shape,
and must explain why the operation is required. See
[VBA security scanning](docs/developer-guide/scan-vba-security.md).

Editing VBA invalidates an existing digital signature. Build, functional
testing, and signing belong to the release process after the reviewed source
has been approved.

## Change Python code

- Add or update tests for changed behaviour.
- Let Black and isort format the code through pre-commit.
- Avoid hidden dependencies on a local Excel installation unless the tool is
  explicitly Windows/Excel-specific.
- Do not make tests depend on network access, production databases, or local
  credentials.

## Change documentation

Documentation is reviewed through pull requests like source code. Update it in
the same pull request whenever behaviour, configuration, security controls, or
supported workflows change.

The documentation project is a Quarto book below `docs/`. Quarto can preview
the HTML book and render a consolidated Word document:

```powershell
quarto preview docs
quarto render docs --to html
quarto render docs --to docx
```

Use the exact names found in the source code for modules, classes, forms,
commands, tables, and settings. Do not include realistic production data in
examples or screenshots. Prefer text-based Mermaid diagrams where practical so
architecture changes remain reviewable in Git.

## Open a pull request

Complete the pull request template and include:

- the problem and the scope of the solution;
- a linked issue when one exists;
- automated and manual test evidence;
- security, database, compatibility, and release impact;
- documentation and changelog impact;
- a rollback or recovery note when the change can affect persisted data or a
  release artifact.

Small follow-up work may be deferred only when it has a separate issue with a
clear impact and priority. Do not defer findings that can cause data loss,
corrupt import/export, expose secrets, invalidate a required security control,
or leave a required test failing without explicit maintainer approval.

All required checks and reviews must pass before merge. Do not merge your own
change solely because automation is green; Excel/VBA changes may also require a
documented functional review.
