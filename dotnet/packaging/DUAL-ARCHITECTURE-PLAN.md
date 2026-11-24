# Dual-architecture review and implementation plan

Source review: 2026-11-13. Requirements: x86 and x64 Excel, shared functionality,
normal startup and updates without administrator credentials. This is a static
review of the delivery path and relevant VBA interop, not certification of all
workbooks or dependencies. Existing x64 artifacts remain unchanged.

## Findings and change map

P1 means resolve before declaring dual-architecture support; P2 means required
supporting verification or design work. Intentional prototype restrictions are
not regressions. Paths below are relative to the repository root.

| ID | Priority | Evidence / location | Required action |
| --- | --- | --- | --- |
| A1 | P1 | `dotnet/src/Nadabas.ExcelBootstrap/Nadabas.ExcelBootstrap.csproj`: `ExcelDnaCreate32BitAddIn=false`; references modern PackageUpdate and VersionStore | Produce both XLLs from shared source. Keep the approved Framework startup prototype separate from modern dependencies; do not merely retarget this project to net48. Inspect native PE architecture of outputs. |
| A2 | P1 | `tools/build_loader_test_installer.ps1`: single SignedLoader/SignedPackageUpdater inputs; VersionStore published with `win-x64` and `AppHostDotNetSearch=Global` | Introduce explicit component architecture inputs and build records. Replace the global-only runtime assumption when the private-runtime contract is proven. Select helper architecture deliberately, separately from Excel. |
| A3 | P1 | `dotnet/packaging/loader-lifecycle.iss`: `PrerequisiteError` reads one App Paths view and rejects binary kind other than 6; `runtime-prerequisite.iss` only checks x64 shared runtime | Discover and inspect Excel before choosing components; reject ambiguous/unsupported targets without mutation. Add private-runtime validation for x86 and x64. Keep current guards until replacement logic is tested. |
| A4 | P1 | `dotnet/packaging/loader-test.iss`: fixed loader path, `ArchitecturesAllowed=x64compatible`; lifecycle and `legacy-handoff.iss` use fixed registry views | Route the selected payload and registration consistently. The architecture allowance is an OS restriction, not an Office detector: x86 Excel on x64 Windows already passes that OS check. Review actual registry sharing before changing views. |
| A5 | P1 | `dotnet/src/Nadabas.PackageUpdate/SignedFeed.cs`: FeedRelease has no architecture, Select filters OS/version only and rejects duplicate selected versions; `dotnet/tools/Nadabas.FeedSigning/Program.cs` emits a Windows enclosure | Do not publish two architecture enclosures at the same version into the current feed. Prefer one compatible universal package for old clients, or introduce an explicitly versioned selection contract with a bridge release. Keep signatures, pins and URL bounds. |
| A6 | P1 | `dotnet/src/Nadabas.PackageUpdate/PackageJob.cs`: PendingPackage schema 1 lacks architecture/runtime identity; VerifyPackage checks publisher/product/version, not payload compatibility | Define installed Excel and payload identity, recheck before apply, and reject stale jobs after an Office architecture change. Preserve legacy schema/old-client behavior explicitly. Setup EXE architecture alone does not identify the contained XLL. |
| A7 | P1 | `PackageJob.Stage` copies only the updater EXE to external cache; `PlatformHandoff.LaunchWaiter`, `AutomaticUpdates.CheckFromExcel`, `ExcelBootstrap/RestartUpdate.cs`, `VersionStore/ExcelRestart.cs` launch helpers/Excel without a private-runtime contract | Make all child processes and cached recovery helpers resolve a verified runtime. Keep the previous usable runtime while a helper is running; do not replace/delete it during upgrade. Preserve save prompts and PID/start-time validation. Test helper/Excel architecture combinations. |
| A8 | P1 | `vba/modules/InterfaceDropClose.bas`: VBA7 declarations and local hWndForm/hMenu use Long for HWND/HMENU; many forms call DropClose | Correct handle parameters, return values and locals to LongPtr in the VBA7 branch; retain 32-bit BOOL/UINT values and the legacy branch. This is an existing x64 correctness defect, not an x86-only blocker. Rebuild and re-sign a new VBA artifact after acceptance; never alter released XLAM bytes. |
| A9 | P2 | `vba/modules/cmdOpen_Close_Database.bas`: ACE and SQLOLEDB connections; `cmdBackUp.bas`: ADOX/ACE creation | Validate provider loading inside each Excel architecture on test databases. Do not silently install or replace Office/ACE drivers. Record missing-provider diagnostics and an IT route; runtime success alone does not prove database compatibility. |
| A10 | P2 | `.github/workflows/dotnet.yml`: OS matrix only; `tools/test_version_loader.ps1`, `test_local_runtime.ps1`, `test_excel_local_runtime.ps1`, `test_legacy_migration.ps1` contain x64 assumptions | Add architecture build checks and separate desktop evidence. Parameterize probes without weakening publisher checks or changing the maintainer's Office. Audit test-runner bitness independently of Excel bitness. |

## Reuse and avoid unnecessary changes

- `Nadabas.Setup/Diagnostics.cs` already reads both registry views and explicitly
  says registration is not proof of Office bitness. Extend this representation
  with verified PE architecture rather than create unrelated detection logic.
- `Nadabas.Setup/LegacyMigration.cs` and `TestAddinRegistration.cs` already inspect
  both views. Reconcile these with installer behavior and deduplicate shared keys.
- `PackageUpdate/UpdatePolicy.cs` deliberately reads machine policy from Registry64.
  On x64 Windows, preserve that documented policy for x86 clients too. Changing it
  to process-default view could cause IT restrictions to be missed. Registry64 in
  `PackageJob.InstalledVersion` needs compatibility tests, not a blind substitution.
- `PublisherTrust.cs` uses IntPtr for native pointers. Retain the verification
  mechanism and test actual struct layouts and valid/invalid signatures in x86.
- `InterfaceShellExec.bas` and `InterfaceWEB.bas` already use LongPtr in VBA7.
  GetUserName buffer lengths, timeGetTime, and clsTreeView POINTAPI coordinates
  remain 32-bit values; do not mechanically replace every Long.
- `Nadabas.UpdateProbe` and `build_installer_prototype.ps1` are older prototype
  paths. Inventory them, but do not refactor them unless they remain in the chosen
  release path. No duplicate runtime management or updater engines.

## Small, reviewable implementation sequence

### 1. Read-only discovery and compatibility contract

Implement a reusable result with Excel path, verified x86/x64 architecture,
discovery evidence, OS support and ambiguity status. Separate Excel architecture
from helper architecture and policy registry view. Agree on supported Windows /
Office versions and available x86 test environment before claiming coverage.

Acceptance: PE fixtures for x86/x64, unsupported machine, malformed/truncated
files, missing executable, duplicate registry paths and ambiguous installations.
No writes to Office registration. This is the recommended first code change.

### 2. Shared build, two startup variants

Build the approved thin net48 startup prototype for x86 and x64. Build modern
XLLs from the same source and isolate output directories. Keep x64 production
artifacts intact. Verify the private-runtime host selection occurs before modern
code loads and does not leave permanent or unrelated process environment changes.
Handle other add-ins already using a modern runtime explicitly; no forced unload.

Acceptance: successful builds and PE checks for both; standard-user normal Excel
startup, no special shortcut, with and without a shared runtime. Unsigned build
evidence is separate from Office acceptance, which needs signed final components.

### 3. VBA interoperability and provider regression

Fix A8 in exported source, preserving CRLF and form resources. Compile/import into
test copies for both architectures. Exercise About and representative forms using
DropClose, then load/save, backup and batch against identified test databases.
Measure a representative large workbook in x86: memory limits cannot be assumed
equivalent to x64. Keep calculations and persisted data formats unchanged.

### 4. Runtime lifecycle and installer selection

Implement the validated runtime acquisition/layout and selection contract, then
wire installer detection to it. Use verified, versioned runtime directories and
bounded downloads. Cover offline/IT deployment, cancellation, disk exhaustion,
partial download, retry and wrong-architecture payloads. Preserve existing runtime
and active-version pointers until the replacement is verified. Uninstall removes
only owned files; never shared machine runtimes or user databases.

Acceptance: fresh standard-user installs on clean x86-Excel and x64-Excel images;
all four languages, normal startup and reboot. Existing 0.9.x x64 upgrade remains
valid. Do not remove current prerequisite guards earlier.

### 5. Update, restart and recovery compatibility

Choose universal signed package vs a versioned architecture-aware feed before
changing A5/A6. Preserve product identity expected by installed clients. Exercise
old jobs, old feeds, signed wrong-architecture packages, IT pins and Office bitness
changes between stage and apply. Preserve existing anti-replay and recovery rules.
Recheck the 128 MiB package limit against measured bundle sizes; do not remove
download bounds simply to fit both runtimes.

Acceptance: extend SignedFeedTests, PackageRulesTests, PackageRestartTests and
ExcelRestartTests; run test_package_update.ps1, test_platform_restart.ps1 and
legacy migration checks for each supported arrangement. Restart and repair must
work when the helper runs outside the installation directory. A rollback cannot
activate an XLL incompatible with the newly installed Office architecture.

### 6. Release gates

Run required repository checks plus the desktop matrix below. Review source and
artifact identities, then perform deferred GlobalSign signing and final signed
acceptance. Public publication remains separately authorized. Do not mark x86
supported based only on a successful build or an x64 simulation.

| Environment | Required evidence |
| --- | --- |
| Supported x64 Windows + x86 Excel | Fresh standard-user install without modern runtime; startup, language, database, update/restart, recovery and uninstall |
| Supported x64 Windows + x64 Excel | Same matrix plus upgrade from the existing x64 package |
| Both, shared runtime already present | Coexistence, correct runtime selection and no modification of shared runtime |
| Office architecture changed after installation/staging | Safe refusal or explicit repair; no mismatched activation or data loss |
| Both, missing provider / offline / denied software policy | Accurate diagnostics and preserved installation/data; no security bypass |

## Evidence and open decisions

Step 2 progress: the [Framework startup prototype](FRAMEWORK-STARTUP-PROTOTYPE.md)
is implemented as a separate net48 project. Both startup and modern bootstrap
now emit x86 and x64 XLLs. A1's disabled-x86 setting above records the initial
review state and is no longer current. The prototype defaults to refusing
activation without approved embedded hashes; it is not integrated into Setup.
Signed desktop acceptance and the helper runtime lifecycle remain open.

Step 1 progress, 2026-11-13: `Nadabas.Setup/ExcelDiscovery.cs` now exposes
read-only PE architecture inspection through the existing `diagnose` command.
Distinct registrations (including stale paths) remain ambiguous; duplicate paths
retain their registry evidence. Invalid, unsupported and missing images cannot
be selected, and registry errors suppress selection. Sixteen new regression
cases cover x86/x64 PE fixtures and failure paths. All 153 .NET tests passed;
the real `diagnose` entry point identified this workstation's x64 Excel without
starting it. This does not test x86 Excel execution or enable installer support.
OS compatibility remains explicitly unverified, pending the deployment matrix.

The initial static review changed documentation only; implementation progress is
recorded above. No new x86 runtime build, Office execution test, database write,
signing or publication was performed in step 1. Prior isolated x64 runtime
acceptance does not satisfy this matrix. The startup prototype remains pending.

Open deployment inputs: supported Windows/Office versions; separate x86 Excel
test machine or disposable VM; representative Access and SQL Server test fixtures;
offline packaging needs. None blocks the read-only discovery implementation.

References: [Microsoft VBA pointer guidance](https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/64-bit-visual-basic-for-applications-overview),
[WOW64 shared/redirected keys](https://learn.microsoft.com/en-us/windows/win32/winprog64/shared-registry-keys),
[distribution requirements](EXCEL-ARCHITECTURE.md).
