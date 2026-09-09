# NADABAS

**National Accounts Database System**

[![Tests](https://github.com/statisticsnorway/nadabas/actions/workflows/tests.yml/badge.svg?branch=main)](https://github.com/statisticsnorway/nadabas/actions/workflows/tests.yml)
[![VBA security scan](https://github.com/statisticsnorway/nadabas/actions/workflows/security.yml/badge.svg?branch=main)](https://github.com/statisticsnorway/nadabas/actions/workflows/security.yml)
[![Documentation](https://github.com/statisticsnorway/nadabas/actions/workflows/documentation.yml/badge.svg?branch=main)](https://github.com/statisticsnorway/nadabas/actions/workflows/documentation.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE.md)

NADABAS is an Excel-first system for structuring national accounts compilation.
It keeps calculations visible in Microsoft Excel while replacing fragile direct
links between workbooks with controlled transfers through a shared database.

This repository contains the version-controlled VBA source, supporting Python
tools, automated checks, and technical documentation. The Excel add-in is a
build and release artifact; the exported files below [`vba/`](vba/) are the
reviewable source of truth.

- **Project website:** <https://sites.google.com/nadabas.net/nadabas>
- **Technical documentation:** <https://statisticsnorway.github.io/nadabas/>
- **Training and user resources:**
  <https://sites.google.com/nadabas.net/nadabas/nadabas-resources>

## Quick navigation

### I want to use or learn NADABAS

- Download and install NADABAS from the
  [documents and downloads page](https://nadabas.net/nadabas/documents-and-downloads).
- For release 6.02.002, keep the two extracted, signed files together and start
  the installation from `NADABAS.6.02.002.xlsm`. The companion
  `NADABAS.xlam` is the add-in that contains the VBA project. See the
  [6.02.002 release and installation record](docs/operations/release-6.02.002.qmd)
  for checksums, signature checks, and the clean-install procedure.
- Read the [NADABAS overview](https://sites.google.com/nadabas.net/nadabas/overview).
- Follow the [tutorials and training resources](https://sites.google.com/nadabas.net/nadabas/nadabas-resources)
  for installation, database creation, classifications, key families, workbook
  registration, data transfers, and batch processing.
- Use the [project contact and registration page](https://sites.google.com/nadabas.net/nadabas/contact-and-registration)
  for training-package or access questions.

### I maintain a NADABAS installation

- Start with the [technical documentation](https://statisticsnorway.github.io/nadabas/).
- Review the [system architecture](docs/technical/architecture.qmd) and
  [VBA internals](docs/vba/index.qmd).
- Check the [database access and schema guide](docs/vba/database-access.qmd)
  before changing database structures or configuration.
- Treat backup, restore, conversion, credentials, and multi-user behaviour as
  deployment-specific operations that require testing with non-production data.

### I want to develop or contribute

- Follow [Getting started](docs/developer-guide/getting-started.md).
- Read the [contribution guide](CONTRIBUTING.md) before making changes.
- Use the [VBA export](docs/developer-guide/export-vba.md) and
  [VBA import](docs/developer-guide/import-vba.md) workflows.
- Review [VBA security scanning](docs/developer-guide/scan-vba-security.md)
  before approving native APIs, file operations, shell access, or process control.

## What NADABAS does

NADABAS connects registered Excel workbooks to a Microsoft Access or SQL Server
database through an Excel ribbon and VBA workflows.

Key capabilities include:

- loading data from a database into defined Excel ranges;
- saving workbook values and provenance back to the database;
- organizing data through key families and dimensions;
- managing classifications and correspondences;
- registering, opening, reserving, and coordinating workbooks;
- running ordered batches of workbook updates;
- linking supporting documents and generating metadata reports;
- providing administrator, permission, period, backup, and diagnostic tools;
- loading translated ribbon, form, message, and error text from workbook resources.

![Before NADABAS, workbooks exchange data through many direct links. With
NADABAS, workbooks exchange data through one shared
database.](docs/assets/nadabas-before-after.svg)

Before NADABAS, workbook-to-workbook links create a fragile network. With
NADABAS, workbooks read and write through a shared database while
classifications, correspondences, permissions, and provenance are managed
centrally.

NADABAS currently depends on desktop Microsoft Excel, VBA, Windows APIs, and
database drivers. The repository's Python tools and documentation checks do not
replace functional testing in a supported Excel environment. See the
[architecture evidence register](docs/technical/architecture-evidence.qmd) for
the distinction between source-verified and environment-dependent behaviour.

## Repository layout

```text
vba/                     Exported VBA source
  modules/               Standard modules (.bas)
  classes/               Class and workbook modules (.cls)
  forms/                 UserForms and binary resources (.frm/.frx)
  customUI/              Ribbon XML, relationships, and image resources
  resources/             Reviewed additions to hidden language worksheets
tools/                   VBA export, import, and security-scanning tools
tests/                   Automated tests and test fixtures
docs/                    Quarto technical documentation and contributor guides
.github/workflows/       Tests, security checks, and documentation deployment
experimental/            Isolated exploratory work; not production source
```

VBA text files use CRLF line endings because Excel can otherwise import some
components incorrectly. Keep every UserForm `.frm` file together with its
required `.frx` resource.

## Development setup

### Prerequisites

- Git;
- Python 3.12 or newer;
- Poetry 2.2 or newer;
- pre-commit;
- Windows and Microsoft Excel for VBA import/export and functional testing;
- Quarto for local documentation preview.

Clone the repository and install the development environment:

```powershell
git clone https://github.com/statisticsnorway/nadabas.git
Set-Location nadabas
poetry install
poetry run pre-commit install
```

Run the same core checks used by GitHub Actions:

```powershell
poetry run pre-commit run --all-files
poetry run pytest -v
poetry run python tools/scan_vba_security.py vba `
  --approvals tools/vba-security-approvals.json `
  --show-approved
```

The security scanner reports operations that require human review. A warning is
not automatically a vulnerability, and a documented approval is not a general
exception for future code.

## VBA source workflow

1. Export a complete VBA project to a staging directory with
   [`tools/export_vba.py`](tools/export_vba.py).
2. Review the staged differences before copying intended changes into `vba/`.
3. Run pre-commit, tests, and the VBA security scanner.
4. Dry-run the import with [`tools/import_vba.py`](tools/import_vba.py).
5. Import into a copy of the approved Excel template and perform documented
   functional tests.
6. Build and sign release artifacts only after source review and approval.

Never use an `.xlam` or `.xlsm` file to silently overwrite reviewed source.
Editing VBA invalidates an existing digital signature.

## Documentation

The technical documentation is a Quarto book published through GitHub Pages.
Preview or render it locally from the repository root:

```powershell
quarto preview docs
quarto render docs --to html
```

Documentation changes should accompany changes to behaviour, configuration,
security controls, or supported workflows. Exact module, class, form, table,
and setting names should match the exported source.

## Project status

The repository is actively establishing reproducible source control, automated
review, and technical documentation for an existing Excel/VBA system. Static
checks cover the exported source and tooling, but Excel behaviour, supported
Office/database combinations, signing, and deployment-specific recovery still
require manual evidence.

For project history, governance, and adoption information, see the official
[background](https://sites.google.com/nadabas.net/nadabas/background) and
[countries](https://sites.google.com/nadabas.net/nadabas/countries) pages.

## Contributing

Contributions are welcome. Keep changes focused, avoid production data and
credentials, and include automated and manual verification appropriate to the
risk. See [`CONTRIBUTING.md`](CONTRIBUTING.md) for the complete workflow.

## Security

Do not report suspected vulnerabilities in a public issue. Follow
[`SECURITY.md`](SECURITY.md) and use GitHub private vulnerability reporting.

## License

This repository is published under the [MIT License](LICENSE.md). Individual
third-party components may retain additional notices or distribution terms;
see the [component reference](docs/vba/localization-and-components.qmd#third-party-tree-view-classes)
before redistributing extracted VBA modules.
