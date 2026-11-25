# Private installer integration review — 2026-11-23

Scope: installation and migration of existing NADABAS installations, without
GlobalSign signing, normal Office registration changes or public publication.
Source branch: `migrating-begins`; the working tree includes uncommitted work.

## Implemented

- One installer carries x86/x64 components from shared source and selects the
  payload by the registered Excel executable's PE architecture.
- The selected modern .NET runtime is private to the per-user installation.
  Framework 4.8 remains an IT-managed prerequisite. Setup does not elevate or
  set permanent runtime environment variables.
- The approved VBA 6.02.009 slot is activated through the selected private
  runtime. First installation exposes language selection; silent installations
  require a language and updates preserve the preference.
- Existing managed-loader registrations are changed in their original OPEN
  slot. Approved legacy 6.02.005 bytes and registration are backed up before
  handoff. Uninstall restores the legacy registration and retains recovery files.
- Modified legacy files, conflicting registrations, architecture changes and
  unreviewed runtime patch transitions are refused before replacement.
- Unsigned normal-profile packages refuse installation. Only the isolated smoke
  identity may install unsigned components; it does not activate them in Excel.

## Evidence

Local reports are under ignored `artifacts/private-installer/`; keep these with
the build records rather than committing generated installers or machine paths.

- Real x86 and x64 private-runtime helper execution through the Inno lifecycle:
  38 checks across fresh install, registration, approved VBA selection, repair,
  language retention/change, old-loader transition, legacy backup, modified-file
  rejection, runtime-patch rejection, uninstall and user-data sentinel retention.
- Automatic discovery of the installed x64 Excel executable without an
  architecture override, English selection and fresh uninstall passed.
- The shared-runtime installer profile still compiles.
- Python suite: 137 passed, one skipped (symlink permission fixture).
- Changed-file pre-commit hooks passed. VBA security scan: zero errors, 14
  warnings, 23 informational findings, two approved findings. VBA was not edited.
- Compressed dual-architecture preparation is about 102 MiB, below the existing
  128 MiB updater bound. Final bytes/hashes belong in build records.

The final review corrected silent-language validation so it reaches the
preflight error instead of opening a modal dialog. Keep the targeted silent-mode
and unsigned-profile refusal reports with the final rebuilt artifacts.

## Remaining acceptance and scope limits

These checks are engineering acceptance, not certification of a distributable
1.0 release. Final signing and ordinary-user Excel startup/update/restart remain
required, including a real 32-bit Excel environment. No such environment is
available on the current workstation.

Legacy migration currently covers the reviewed 6.02.005 identity at the standard
per-user AddIns location and the existing managed install root. Other versions,
custom locations and conflicting registrations require review rather than
automatic overwrite.

Automatic retirement and recovery of a different private .NET patch is not yet
implemented. The installer fails closed for that transition. This must be
completed before claiming full long-term runtime servicing. Also retain the
runtime provenance/redistribution review and signed-startup coexistence gates in
[PRIVATE-RUNTIME-BUILD.md](PRIVATE-RUNTIME-BUILD.md).

The migration restoration tests do not establish automatic package rollback for
every crash or power loss. Preserve the existing explicit recovery procedures.
