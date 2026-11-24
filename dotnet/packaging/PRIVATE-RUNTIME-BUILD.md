# Private runtime component preparation

This engineering preparation builds both architectures from shared source. The component
preparation does not register Excel add-ins, sign files or publish a release.
The separate unsigned installer preparation below integrates these components
with the shared Inno lifecycle. Do not substitute files into a signed installer.

## Inputs and command

Provide one plain runtime directory per architecture, containing `dotnet.exe`,
`host/fxr/<version>` and matching `shared/Microsoft.NETCore.App/<version>` and
`shared/Microsoft.WindowsDesktop.App/<version>` trees. Use the same stable .NET
10 patch version for both. The preparation checks native host/Core architecture,
complete required entries, file hashes and matching versions. These checks do
not establish Microsoft provenance or redistribution approval. Review original
runtime sources, servicing policy and redistribution files before release.

Use a fresh short path to avoid Windows packaging path limits:

```powershell
$build = Join-Path $env:TEMP ('nb-private-' + [guid]::NewGuid().ToString('N'))
./tools/prepare_private_runtime_build.ps1 `
  -RuntimeX86 <absolute-x86-runtime-root> `
  -RuntimeX64 <absolute-x64-runtime-root> `
  -Output $build
```

The script uses the repository's Python environment and installed .NET SDK. It
never overwrites an existing output directory. A failed build retains its output
for diagnosis and does not write a successful component record. Python runtime
preparation can also be exercised separately through
`tools/prepare_runtime_bundle.py`; tests are in `tests/test_runtime_bundle.py`.

## Output contract

`runtime-manifest.tsv` contains the combined architecture/path/SHA-256 input for
PackageUpdate. Both updater builds and the modern loader use this same manifest.
`bundle-record.json` records runtime versions, counts and uncompressed sizes.

Each `payloads/x86` or `payloads/x64` directory contains the selected runtime
under `runtime/<arch>`, the matching modern loader, Framework startup XLL,
updater and VersionStore tools. Helpers are published with environment-only
runtime discovery. `component-build.json` records component hashes and source
state only after all builds succeed. Both records explicitly forbid distribution.
The installer preparation adds the approved XLAM slots and startup registration.

`startup-manifest.UNSIGNED.tsv` is an engineering manifest bound to unsigned
modern-loader bytes. The startup XLL still enforces publisher verification, so
this output is not an Office-trust workaround and is not desktop acceptance.

## Release sequence still required

1. Approve runtime provenance, redistribution files and servicing policy.
2. Build and sign the modern loaders and helpers with the reviewed runtime
   manifest. Verify final signatures and architecture.
3. Regenerate startup hashes using final signed modern-loader bytes, rebuild
   the Framework startup with that manifest, then sign and verify startup.
4. Build the installer with the approved XLAM store and the final signed components.
   The current preparation script deliberately produces unsigned artifacts only.
5. Measure the compressed package against the existing update-size bound. Run
   signed installation, restart, recovery and uninstall acceptance. Actual
   32-bit Excel acceptance remains pending a suitable environment.

See [the prototype](FRAMEWORK-STARTUP-PROTOTYPE.md) and
[the architecture plan](DUAL-ARCHITECTURE-PLAN.md).

## Unsigned installer integration

`tools/build_private_runtime_installer.ps1` accepts the component build, an earlier
build record (immutable historical slot hashes), approved baseline and 6.02.009
XLAMs, the Inno compiler, and a new absolute output directory. `-SmokeTest` routes
all Excel and language registration into `Software\NADABAS\InstallerSmoke`.
Without that switch the unsigned package refuses installation before changes.
No token or production signing key is accessed.

The `PrivateRuntime` profile shares the existing installer, legacy migration,
language selection and version-store logic. It carries both architectures and
copies only the architecture identified in Excel's PE header, consulting both
machine App Paths registry views. Missing, unsupported or conflicting entries
stop installation. Framework 4.8 must already be enabled by IT; modern .NET is
bundled privately. Setup stays per-user, without elevation or permanent runtime
environment variables. Activation invokes the selected private `dotnet.exe`
explicitly. EN/FR/PT/ID selection remains visible at first interactive install;
silent first installs require `/NADABASLANG`, and updates preserve the selection.

The recorded Excel architecture is also used for uninstall. An architecture
change is refused before replacement and requires uninstall/reinstall. It does
not silently move a legacy journal between registry views.

### Existing installations and recovery

An existing managed loader at the same install root is replaced in its original
OPEN slot, preserving migration journal compatibility. The approved legacy
6.02.005 add-in in the user's AddIns directory is copied and hash-verified before
registration handoff. The original file remains in place. Uninstall restores its
original startup registration, retains the journal and backup, and does not
remove user workbooks, databases or language settings.

Unknown NADABAS startup paths, modified legacy bytes, conflicting startup entries,
filesystem links and missing/conflicting recovery journals stop for review.
These cases are not silently overwritten. Legacy versions other than the
approved baseline and custom locations still require a reviewed migration path.
This is explicit legacy restoration, not a claim of automatic package rollback
following every interrupted installation.

### Acceptance

Run `tools/test_private_runtime_installer.ps1` with the smoke build directory and
approved baseline. It executes real x86/x64 private .NET helpers and the shared
installer engine against synthetic registration. The architecture override exists
only in SmokeTest builds; it does not establish actual x86 Excel acceptance.
Keep its report and Setup logs. Signed normal-user Excel startup, restart,
coexistence with other .NET add-ins, actual 32-bit Excel and GitHub distribution
remain separate release gates.

### Runtime patch retirement gate

The current profile supports installation of its reviewed runtime and package
updates using that same runtime patch. It rejects an existing different patch
before replacement: retaining two patches would violate startup's exact runtime
inventory, while deleting unreviewed files would be unsafe. Automatic runtime
patch transitions need a reviewed retirement/recovery transaction before this
profile can be considered complete for long-term production servicing. Do not
remove this guard merely to make a future package install.

## Signed internal candidates

After preparing a fresh component build, `tools/sign_private_runtime_components.ps1`
accepts its directory, the local SignTool path and the approved publisher thumbprint.
It verifies the unsigned input hashes, signs and verifies the helpers, managed
VersionStore entry points and modern loaders, then regenerates the startup manifest
from the final signed loader hashes. It rebuilds and signs both Framework startup
XLLs. Incomplete signing never produces a successful signed component record.
PIN entry remains in the token's own window.

`tools/build_private_runtime_installer.ps1` accepts `-SignTool` and
`-CertificateThumbprint` for this signed component record. It verifies runtime
inventory, component hashes and publisher signatures before compiling. The signed
startup manifest must pin the final signed loaders. Setup signs its uninstaller
and final installer, and the build record still marks the artifact internal and
not approved for distribution.

Optional `-AccessFixAddin` stages the exact signed 6.02.010 candidate in a new
`slot-fix602010`, preserving the existing 6.02.009 slot and identity. Setup selects
the new slot when that build input is present; older build profiles retain their
previous selection. The signed candidate is recorded in
`../releases/access-repair-6.02.010.json`. Final ordinary Excel database acceptance,
source approval and public distribution are separate from signing.

The maintainer has accepted deferring real x86 Excel acceptance to a colleague.
Record that gap explicitly; isolated x86 helper execution is not an x86 Excel test.

## Agreed distribution direction — 2026-11-24

The maintainer approved a small default online installer instead of distributing
the complete dual-architecture runtime with every NADABAS update. Detect Excel
architecture, download and verify only its required runtime when absent, install
it per-user, and reuse the verified runtime for later program-only updates.
Retain the complete offline package as an optional artifact for IT and offline
use. First installation still requires runtime bytes if no suitable runtime is
available; this direction does not claim that dependency has disappeared.

This split is not implemented yet. The current signed candidate is still the
complete offline package. Add authenticated download, interrupted-download retry,
offline behavior and safe runtime patch servicing before presenting the smaller
installer as ready. Hash and publisher checks remain mandatory.
