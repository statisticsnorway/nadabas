# Excel architecture and distribution

See the [code review and staged implementation plan](DUAL-ARCHITECTURE-PLAN.md)
for concrete change locations, compatibility risks and acceptance gates.

## Requirements and recommendation

The maintainer confirmed on 2026-11-13 that both x86 and x64 Excel are required.
The standard-user installation requirement applies to both. These requirements
do not imply support for every Windows version, 32-bit Windows, or native ARM64
Office; those platform boundaries need a separate support matrix.

Recommended distribution: one user-facing Setup entry point, selecting the
correct architecture automatically. Produce separate signed x86 and x64 payloads
from the same source, version and release process. Do not create two products or
two business-logic branches. IT may receive architecture-specific offline packages
from the same build pipeline. The exact online/offline packaging format remains
to be implemented and tested.

An online bootstrap can keep the initial download small by obtaining only the
required verified runtime and payload. An offline bundle avoids network dependency
but is larger. Both must use the same payloads and verification rules; neither
may require the end user to select a bitness. Do not download or execute an
unverified dependency. Preserve the existing SSH, feed and publisher checks.

## Component boundaries

| Component | Architecture rule |
| --- | --- |
| Excel startup component / XLL | x86 for x86 Excel, x64 for x64 Excel |
| Modern runtime loaded inside Excel | Match Excel, even on 64-bit Windows |
| Shared managed business logic | One source; isolate native interoperability |
| VBA and workbooks | Shared where compatible; audit declarations and providers |
| Native database provider loaded inside Excel | Match Excel and the supported backend |
| Separate updater/helper process | Choose a supported OS architecture explicitly; do not assume it must match Excel |

The approved next experiment is a small .NET Framework 4.8 startup component
that locates a private modern runtime before loading the existing modern XLL.
Build this experiment for both Excel architectures. This is not a decision to
port the business logic back to Framework. A thin bridge must retain publisher
verification and must not change permanent runtime selection for unrelated apps.
Framework availability itself must be checked on supported Windows systems.

## Installer and update contract

1. Discover Excel, inspect its executable architecture and resolve ambiguity
   before changing files. Do not infer Office bitness from Windows bitness.
2. Select compatible startup, XLL and runtime artifacts. Verify their identity
   and signatures before activation. Keep the language choice and user-facing
   experience the same for both architectures.
3. Include architecture in release selection and diagnostics. If using a universal
   installer, validate its selected payload too. Distinct artifact bytes need
   distinct immutable identities/hashes; preserve compatibility with old clients.
4. Recheck Excel architecture during update. An Office x86/x64 change requires a
   controlled repair or migration with Excel closed, not an ordinary blind update.
   Do not roll back an incompatible XLL after Office has changed architecture.
5. Keep workbooks, databases, preferences and recognized old installations intact.
   Missing database drivers must be reported separately from missing runtime.

## Current gaps and acceptance

Both Framework startup and modern ExcelBootstrap now build x86 and x64 XLLs.
The installer still requires 64-bit Excel, VersionStore publishing uses win-x64,
and the local-runtime probes use x64 components. Builds do not establish desktop
support for x86. Removing the installer guard alone would allow an unusable
installation. See [prototype limitations](FRAMEWORK-STARTUP-PROTOTYPE.md).

Use separate x86-Excel and x64-Excel environments, including x86 Excel on x64
Windows. In each, test standard-user install without modern .NET, normal Excel
startup, approved VBA loading, all supported languages, representative database
load/save, update/restart, cancellation, recovery and uninstall. Exercise incorrect
architecture, missing providers and Office architecture changes. Do not replace
the maintainer's Office installation to obtain coverage.

The isolated x64 local-runtime test established that the existing signed loader
can load VBA 6.02.009 using private .NET. It does not establish normal startup,
x86 support or acceptance on a clean machine. See [local runtime](LOCAL-RUNTIME.md).

Excel-DNA supports packed add-ins for both architectures and requires the modern
runtime architecture to match Excel:
[installation](https://excel-dna.net/docs/guides-basic/installing-your-add-in/),
[runtime support](https://excel-dna.net/docs/guides-basic/dotnet-runtime-support/).
