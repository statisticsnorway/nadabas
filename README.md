# NADABAS

National Accounts Database System (NADABAS)

NADABAS is a system for compiling national accounts statistics
using Microsoft Excel, VBA, and SQL Server.

## Repository structure

- `vba/` VBA source code exported from Excel
- `sql/` SQL Server schema and migration scripts
- `tools/` Python tools for import/export and build automation
- `docs/` Documentation and installation guides
- `tests/` Automated tests and validation routines

## VBA review and XLAM build

VBA source can be reviewed through normal GitHub pull requests and imported
into a copy of the NADABAS add-in with `tools/import_vba.py`. See
[`docs/developer-guide/import-vba.md`](docs/developer-guide/import-vba.md) for
prerequisites, commands, UserForm handling, and the recommended release flow.

Review security-relevant VBA operations with `tools/scan_vba_security.py`:

```bash
python tools/scan_vba_security.py vba
```

The scanner removes VBA comments, distinguishes errors from review warnings,
and supports narrow, documented approvals. See
[`docs/developer-guide/scan-vba-security.md`](docs/developer-guide/scan-vba-security.md).

## Status

Initial repository setup.
