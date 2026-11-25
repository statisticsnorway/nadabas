# Framework startup prototype

Engineering-only implementation, 2026-11-13. The prototype is not wired into Setup
or the update channel and is not an accepted replacement for the installed XLL.

Both architecture payloads can now be prepared reproducibly through
[private runtime component preparation](PRIVATE-RUNTIME-BUILD.md). This builds
unsigned components and manifests; installer integration and signed acceptance
remain separate gates.

## Implemented boundary

`src/Nadabas.ExcelStartup` targets net48 and builds packed x86 and x64 XLLs from
one source. It references no modern NADABAS project. The existing modern
ExcelBootstrap now also emits both packed architectures; its runtime remains
net10.0-windows. Shared business logic has not been duplicated or retargeted.

On an Excel macro callback, the Framework startup selects by process pointer size,
verifies an embedded file manifest and the modern XLL's SSB Authenticode signature,
then invokes Application.RegisterXLL. The signature verifier is linked from the
existing PackageUpdate source, with framework-specific certificate construction;
both targets use WinVerifyTrust and the same signer extracted from verified state.

Runtime files and the modern XLL are SHA-256 pinned in the startup assembly. The
runtime directory must have exactly the listed files, no filesystem links, and
matching .NET 10 host/Core/Desktop versions. Read leases protect verified files
during activation and remain held by the startup instance. Hashes do not establish
provenance: release operators must approve and verify runtime inputs before
embedding their hashes and signing the startup XLL.

DOTNET_ROOT variables are set only within the Excel process during RegisterXLL
and restored in a finally block. No user/machine environment or registry changes
are made. If CoreCLR is already loaded, the prototype refuses to change runtime
selection. Coexistence with other modern add-ins needs further design and tests.

## Safe default and payload layout

The default embedded `startup-manifest.txt` contains no approved payload. Default
builds intentionally refuse activation. There is no switch to bypass verification.
A reviewed manifest is supplied at build time using `-p:StartupManifest=<path>`.
Each non-comment line contains architecture, relative path and SHA-256 separated
by tabs. Paths use forward slashes. Both architectures can share one manifest.

Expected layout under the existing version-store root:

```text
Nadabas.ExcelStartup-AddIn-packed.xll        # x86 startup, or x64 equivalent
Nadabas.Loader.x86.xll                      # modern, final signed bytes
Nadabas.Loader.x64.xll
runtime/x86/dotnet.exe
runtime/x86/host/fxr/<version>/hostfxr.dll
runtime/x86/shared/Microsoft.NETCore.App/<version>/...
runtime/x86/shared/Microsoft.WindowsDesktop.App/<version>/...
runtime/x64/...                            # same structure for x64
versions/...                              # existing version store
```

Modern XLLs stay at the version-store root because the existing loader derives
its store path from its own XLL location. Do not register both the startup XLL
and the modern XLL independently. Installation and handover are not implemented.

Build without token use from `dotnet/`:

```powershell
dotnet restore src/Nadabas.ExcelStartup --locked-mode
dotnet build src/Nadabas.ExcelStartup -c Release --no-restore
dotnet build src/Nadabas.ExcelBootstrap -c Release --no-restore
```

NuGet supplies reference assemblies for building; this does not install Framework
or Desktop Runtime on a user's PC. Published artifacts from these commands are
unsigned and must not be used to weaken Office trust for testing.

## Remaining acceptance gates

### Helper runtime handoff progress (2026-11-16)

The checker and slot-restart launch paths now derive the runtime root from the
running CLR's directory, check the helper executable architecture and set only
the child's DOTNET_ROOT variables. Shared-runtime launches are explicit too;
the parent environment is untouched. Both helpers can be published with
`-p:PrivateRuntimeHelpers=true`, which selects EnvironmentVariable-only host
discovery. Default builds retain their previous behavior. These new binaries
must be built and signed together with the matching architecture payload.

For engineering verification, publish the updater with that property and
`-r win-x64 --self-contained false -p:PublishSingleFile=true`, then run
`tools/test_private_helper_runtime.ps1 -Updater <exe> -RuntimeDirectory <root>`.
It tests the usage entry point and missing-runtime refusal, not installation.
Run the probe separately for each architecture, passing `-Architecture x86` or
`-Architecture x64` and the matching published updater and runtime directory.
It validates the PE architecture of both the helper and `dotnet.exe` before
launching, and reports architecture and helper hash. Neither run proves Excel
execution, installation, registration or restart/update acceptance.

Verification reviewed on 2026-11-19: the engineering x64 apphost passed both
private-runtime startup and missing-runtime refusal in
`artifacts/private-helper-runtime/test-587ae1eb0789451795becb7730c5a0de/report.json`.
The probe runs the usage entry point only; it does not install or update NADABAS.

Whole-package restart is deliberately blocked when the helper's runtime resides
inside the installation being replaced. RunSetup also enforces this for direct
apply/recovery. The acknowledged launch described below selects the verified
external copy instead. Signed end-to-end acceptance and cache cleanup remain
open; this is not a completed distributable update path. Global installations are not
subject to this new in-directory failure.

Direct apply and recovery reject this arrangement before acquiring the job lock;
apply therefore does not create `apply-started.json` or a repair-required status
for a runtime placement error. RunSetup retains its final guard as well.

### Verified cache staging

`PackageJob.Stage` now stages a private runtime before writing its ready job.
Shared-runtime installations keep the previous path. The runtime-only manifest
is embedded in PackageUpdate using `-p:RuntimeManifest=<reviewed-file>`, with
logical resource name `Nadabas.RuntimeManifest`. Its default contains no approved
files and refuses private-runtime staging. Each TSV line contains architecture,
path relative to that architecture's runtime root, and SHA-256. Pin the exact
host, Core and Desktop files for one matching stable .NET 10 version. This
runtime-only manifest avoids a hash dependency cycle with the modern XLL.

`RuntimeCache` verifies the complete source inventory and hashes under read
leases, copies to a unique pending directory in the external job cache, verifies
the copy, then renames it to `runtime/<architecture>`. A job-specific lock rejects
concurrent staging. A retry may reuse an exact verified final copy; it never
overwrites a damaged copy in place. Partial copies remain inert for diagnosis.
Automatic cleanup is not enabled; explicit maintenance is described below.
Manifest hashes must be reviewed before
the updater is signed; copying does not independently establish provenance.

### Acknowledged runtime handoff

Excel uses `LaunchReadyWaiter`, which opens and verifies the cached runtime,
sets the helper's runtime environment and retains the parent read leases while
waiting for `restart-helper-ready` (up to 30 seconds). The child checks its actual
loaded Core runtime directory against the verified cache and acquires its own
read leases before acknowledging. Only then may Excel quit through its ordinary
save/cancellation flow. The existing unacknowledged `LaunchWaiter` API continues
to refuse private runtimes; existing callers cannot silently skip this handoff.

The child retains its leases while waiting for Excel, applying Setup and
checking the installed version. Direct apply and explicit recovery acquire the
same leases before taking their job locks. A missing, modified or wrong runtime
cannot fall back to shared-runtime behavior. Legacy jobs without a cached
runtime continue to use their shared runtime. Restart failure keeps Excel open;
the launcher terminates only its own failed helper and drains diagnostic output.

Tests cover overlapping parent/child leases, release after refusal, missing and
tampered caches, legacy behavior and missing/incorrect/timed-out acknowledgement.
They use fixture files and tasks, not a signed Excel-to-installer process chain.

### Runtime cleanup and process probe

Runtime leases also hold `runtime-use.lock` for shared reading. Explicit cleanup
requires exclusive installation, apply, stage and runtime-use locks, so it cannot
overlap the normal staging/launch/apply/recovery paths. It accepts only completed
(`installed` or `restored`) jobs, or cancelled jobs that never started applying.
The installation's pending job is always retained. Failed, ambiguous and partial
staging jobs remain available for diagnosis.

The operator may preview and then request removal for one job and architecture:

```text
Nadabas.PackageUpdate.exe plan-runtime-cleanup <installed-root> <job-id> <x86|x64>
Nadabas.PackageUpdate.exe cleanup-runtime <installed-root> <job-id> <x86|x64>
```

Both commands recheck ownership, status, locks, plain paths and the remaining
runtime files against the updater's embedded manifest. Use an updater with the
matching reviewed manifest; unknown versions/files are retained. Preview removes
nothing. Removal deletes only verified files below that job's `runtime/<arch>`
and then empty directories, never recursively. Missing approved files permit a
retry after interrupted removal. Installers, helpers, logs, job records and
`runtime-pending-*` directories are retained. No existing user cache was pruned
as part of engineering validation.

`tools/Nadabas.RuntimeProbe` is a separate engineering executable that links the
actual RuntimeCache/PackageRules source. Publish separately for win-x86 and
win-x64. Run `probe <runtime-source> <report-parent>` with the matching
DOTNET_ROOT architecture variable pointing to the input runtime. It generates
fixture hashes, stages real runtime copies, starts a second process using the
cached runtime and checks that its leases survive parent release and end on
child exit. It removes only its own hash-verified runtime copies afterwards,
retaining the manifest and JSON report. Generated hashes are test inputs, not
release approvals. This probe neither starts Excel nor installs software.

Evidence, 2026-11-20: both real-runtime process probes passed (477 files for x86,
478 for x64). Their generated runtime copies were removed after child exit:

- `artifacts/runtime-handoff-probe/runs/runtime-handoff-59d043cc14044010adbf57352c23d500/report.json`
- `artifacts/runtime-handoff-probe/runs/runtime-handoff-b3b26bcd74d840e389944ba01d38e47f/report.json`

All 210 regression tests passed under both x64 and x86 test hosts. These results
cover the engineering handoff and fixture-based cleanup, not signed installer
acceptance or 32-bit Excel. No 32-bit Excel test environment is currently available.

The remaining work before enabling private-runtime package replacement is:

1. Supply reviewed runtime manifests for the signed build and verify agreement
   with the startup payload's runtime inventory for both architectures.
2. Exercise staging with real runtime files and release inputs in addition to
   the synthetic fixtures. Do not mark a release ready based on fixture tests.
3. Exercise the acknowledged handoff with real signed helper processes through
   installer exit, post-install verification and explicit recovery.
4. Validate the explicit cleanup command with signed release inputs. Automatic
   selection/retention policy is not enabled; failed/interrupted jobs are retained.
5. Test tampering, interrupted copying, parallel launch, cancellation and retry,
   followed by real standard-user x86 and x64 restart/update acceptance.

### Outstanding gates

- Approve runtime sources and generate/review complete manifests for both targets.
  Sign modern payload first, pin its final bytes, then build/sign startup.
- Verify native publisher interop on x86 and net48, not just tests linked into net10.
- Test ordinary Excel startup on both architectures without shared modern .NET.
- Validate the runtime handoff and cleanup with signed release packages.
  Automatic retention policy remains open; cleanup is explicit only.
- Define coexistence, partial RegisterXLL failure and explicit add-in removal.
  Do not unload or replace a runtime inside an active Excel process.
- Add safe installer architecture selection and retain required runtime files
  throughout update, rollback and recovery. Keep current installer guards.
- Run clean standard-user desktop, database and update acceptance before release.

The shared StartupPayload tests validate scope, hashes, missing runtime,
architecture selection, publisher failure cleanup and environment restoration.
They are not Excel acceptance. No signing or publication is authorized by this
document. See [the staged plan](DUAL-ARCHITECTURE-PLAN.md).
