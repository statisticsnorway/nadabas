# Standard-user installation investigation

Accepted direction, 2026-11-12: a normal Windows user must be able to install
NADABAS without administrator credentials, including on machines without the
shared .NET Desktop Runtime. Organization software and Office policies still
apply. A download link requiring administrator installation is not sufficient.

## Current evidence and boundaries

`tools/test_local_runtime.ps1` copies an existing installed runtime into a fresh
ignored artifacts directory. This is an engineering fixture, not a runtime
download or distribution mechanism. It copies no SDK into that runtime. The SDK
provides nethost only for the native-host probe. No Office or machine registry
settings are written, and environment variables are confined to a child process.

The probe uses the native calls used by Excel-DNA 1.9: get_hostfxr_path,
hostfxr_initialize_for_runtime_config, and hostfxr_get_runtime_delegate. It checks
that host and framework resolution stay inside the private runtime directory and
rejects an elevated test process. Native activation does not establish that the
actual Excel add-in or a clean machine passes acceptance.

Run from a normal, non-elevated PowerShell:

```powershell
./tools/test_local_runtime.ps1 -RuntimeVersion <installed-10.0-patch>
```

Runtime copies, traces and reports remain under `artifacts/local-runtime/`.
Do not ship this local copy as a release or alter system installations to mimic
a clean machine. Use a disposable clean Windows/Office environment for that test.

## Acceptance before changing distribution

- Fresh install as a standard user, without SDK or machine-wide modern .NET.
- Normal Excel launch from Start and an existing workbook; no special shortcut.
- Loader, version store, background checker and cached restart helper all resolve
  the intended runtime, including after Windows restart.
- No permanent machine-wide environment, registry, Office security or trust changes.
- Coexistence with other Excel add-ins; process-wide runtime selection must not
  silently break another add-in.
- Download authenticity, interruption, retry and offline/IT-provided installation.
- Runtime security servicing, retained rollback dependencies and ownership-aware
  uninstall. A private runtime is our servicing responsibility.
- Existing shared-runtime installations continue to work.

The current installer still checks the machine-wide runtime. Do not remove this
check until the full local-runtime path is implemented and accepted. The current
guided download page is an interim flow only.

## Open implementation choice

Excel-DNA 1.9 discovers the host before managed NADABAS code runs. Setting an
environment variable inside the managed loader is therefore too late. A test
launcher can supply process-local DOTNET_ROOT, but normal Excel startup cannot
depend on that launcher. Determine a supported early host-selection mechanism
before choosing a local-runtime design. A small Framework loader with modern
helpers outside Excel remains an alternative, not an approved migration yet.

References:
[Excel-DNA 1.9 native host](https://github.com/Excel-DNA/ExcelDna/blob/v1.9.0/Source/ExcelDna.Host/host.cpp),
[Microsoft runtime installation](https://learn.microsoft.com/en-us/dotnet/core/install/windows).
