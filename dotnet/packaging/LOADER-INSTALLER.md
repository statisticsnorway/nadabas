> Source update, 2026-11-11: the next build places **Show Exchange tools**
> immediately after **Create Exchange links** in Administration / Other admin.
> The administrator group and existing workbook-link visibility conditions apply.
> Historical signed packages described below retain their original ribbon.

# Compact test package 0.8.1

VersionStore and the signed single-file updater use the installed x64 .NET 10
runtime. Setup checks for the Desktop Runtime before installation. Publish the
updater into a fresh directory with the following settings before signing it.

    dotnet publish src/Nadabas.PackageUpdate -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -p:DebugType=None -o <new-directory>

Keep the signed executable as the updater input to the installer builder. The
staging command copies that executable into the external update cache. The
loader and VBA versions are the same as in package 0.8.0.

# 0.8 engine candidate

Includes the separately signed package updater in updater/. Build requires
-SignedPackageUpdater and -ContextBetaAddin; the latter is the contextual-toggle
revision in slot-beta602006b, preserving the original beta slot. A build can use
-BuildParent with an absolute local temporary path to avoid sync-folder rename
failures. It creates a new unique directory and never deletes earlier builds.

See PACKAGE-UPDATER-0.8.md for the implemented engine, isolated tests and remaining
UI/recovery work. The current Excel dialog still handles XLAM slot updates only.

# Signed Exchange beta package 0.7.1

Uses the user's actual beta ZIP as input; its version remains 6.02.006. The
compiled allowlist admits a separate slot-beta602006, never the simulated 006
slot. The signed candidate changes only Exchange ribbon visibility, managed
update integration and About build date. Both ribbon schemas gain a persistent
per-user toggle beside About, default off. All five Exchange groups remain
subject to their original workbook/admin conditions after the display gate.
Hiding tools does not disconnect or disable Exchange database functionality.

Build requires -BetaAddin in addition to retained earlier fixtures. Setup selects
the beta explicitly. This tests actual package-content replacement despite the
lower VBA number than simulated 008; automatic queue downgrade protection is
unchanged. It is not the future 0.8 whole-platform automatic-update milestone.
See the included guide for display testing and explicit rollback to 008.

# Signed loader test installer 0.7.0

Adds Restart and update as the primary notification action. It checks unsaved
non-add-in workbooks and other Excel processes, starts a bounded out-of-process
waiter, waits for its readiness handshake, and requests normal Excel Quit with
alerts enabled. Only after that original process exits does the helper recheck
the pending update and other Excel processes and reopen Excel. No process kill
of Excel and no automatic workbook restoration. An aborted shutdown times out
after 60 seconds. Interactive restart/save-prompt acceptance is required.

The simulated 6.02.008 fixture has build date 2026-11-03 stamped before signing.
About reads that fixed metadata. Development source uses today's date marked
(development) only while its date constant is empty. VBA signing does not itself
freeze runtime expressions. `prepare_upgrade_test.py --managed-updates --stamp-date`
stamps the build date automatically, with `--build-date YYYY-MM-DD` for a controlled
release date. The builder requires `-DatedUpgradeAddin` alongside the retained
006 and 007 fixtures. Old signed slots are never rewritten.

# Signed loader test installer 0.6.0

Version 0.6.0 adds a modeless Excel-owned update notification, polling local
verified pending state while Excel is idle. Later suppresses repeated prompts
for that slot in the session; Cancel update removes the pending request.
Reading status revalidates both releases without activating either one.
The Excel command `NADABAS.CheckUpdates` provides the same status on demand.

The isolated 6.02.007 VBA fixture adds a session-only managed-update bridge.
It cancels existing VBA scheduling, prevents database opening from rescheduling
legacy checks, and routes the existing manual ribbon action to the loader.
No saved preference is overwritten. Standalone VBA remains unchanged in behavior;
historical signed 005/006 fixtures lack the bridge and keep their manual action.
GitHub checks/downloads and policy enforcement remain a separate milestone.

Build additionally requires `-ManagedUpgradeAddin` for the signed 007 fixture;
`-UpgradeAddin` remains the immutable 006 fixture. Both hashes are compiled into
the signed loader/CLI allowlist. Setup selects 007; old slots remain available.

## Earlier milestones

### Installer 0.5.0

The 0.5.0 milestone adds pending activation for an already staged, approved VBA
release. `queue <root> <slot>` validates but does not change the active pointer;
the next loader startup revalidates and switches atomically. `cancel-update`
and explicit `activate` cancel pending requests. No network or scheduling is
introduced. The English guide simulates the future automatic updater's readiness
signal with the CLI, then verifies the actual next-start behavior in Excel.

Startup logs now include updateStatus/updateError. A rejected request retains
the previous usable release; a crash after pointer replacement but before pending
cleanup is treated idempotently. Rejected pending records remain for diagnosis
until cancelled; they are not executed or accepted as new trust material.
Uninstall removes the exact pending-update.json state file. Old running sessions
keep their own version lease. This does not update loader binaries while Excel
is running and does not complete the GitHub automatic-update integration.

Version 0.4.0 adds a signed, simulated VBA 6.02.006 fixture to slot-test602006.
Only the version constant and About test label changed from the baseline. The
signed loader carries the approved hash/version catalog and checks the actual
VBA runtime version against the selected release. Setup deliberately selects
this new release after installation; rerunning this package selects it again.
Old slots and reports remain available for rollback. This supersedes the 0.3.2
repair behavior that preserved the active pointer unchanged.

All new installation guidance, messages and shortcuts are English. Preparation
checks for registered .NET 10 Desktop Runtime x64 and 64-bit Excel before changing
the installation. No automatic prerequisite download is included. `-UpgradeAddin`
is now required when building and must be signed by the selected publisher and
match the compiled allowlist. Neither a caller-supplied hash nor a sidecar JSON
file can independently authorize new VBA content.

This is a separate, per-user manual acceptance package for 64-bit Excel. It
installs into `%LOCALAPPDATA%\NADABAS-Loader-Test` with its own Windows uninstall
identity. It updates a recognized test installation in place, preserving the
active pointer, immutable version slots and session reports. Empty destinations
and recognized leftovers (reports, lock file, empty tool directories and
hash-verified baseline XLAM slots) are accepted. Unrecognized directories,
filesystem links and modified baseline slots are rejected before installation.

With Excel closed, Setup adds the signed XLL to the first unused HKCU Office
16.0 Excel OPEN entry, or retains its existing exact entry. Other startup entries
are preserved. Only exact registrations for the two payload XLAM paths inside
this installation are removed when upgrading the old prototype. Uninstall
removes exact loader/payload registrations, never whole Office keys. Setup and
Uninstall check Excel processes via WMI and refuse if Excel is open or the check
fails. No process is forcibly terminated and no Office trust settings change.
Activation occurs on the next Excel start; the installer does not launch Excel.

The package contains the signed Excel-DNA bootstrap, a self-contained
version-store command-line tool, two baseline 6.02.005 slots and the simulated
6.02.006 upgrade slot. The bootstrap still requires the separately installed
.NET 10 Windows Desktop Runtime x64. The CLI runtime does not satisfy that
requirement. No SDK is needed on the test machine.

Build with `tools/build_loader_test_installer.ps1`, passing the signed XLL,
Inno compiler, Windows SignTool and public certificate thumbprint. The builder
verifies the loader's signature and publisher, and the approved XLAM hash.
Inno signs both Setup and Uninstall through the configured token. PIN entry
belongs only in the token UI. The build record contains hashes and public
certificate identity, not credentials. Binaries remain in ignored artifacts.

Inno embeds file data without the source's Zone.Identifier. The signed package
is the distribution trust boundary; there is no recursive or blanket unblock
operation. Actual browser-downloaded installer behavior and organizational
security policies still require separate acceptance testing. Do not infer that
this package removes all Windows or Office security prompts.

The manual guide covers automatic startup, two Excel processes selecting different
slots, next-start selection, rollback, deactivation and uninstall. The baseline
shows 6.02.005 and the simulated release shows 6.02.006 with an explicit test label.
Database read/write and migration rollback are outside this test.

Before broader distribution: validate the manual checks, prerequisite handling,
registration and cleanup in the real Excel profile, actual downloaded-package behavior,
upgrade/recovery, publisher trust, IT policy, and the production update channel.
This package does not include the separate NetSparkle update probe. Installation
and loader tests must pass before integrating unattended updates. Uninstall
requires Excel closure; reports may remain and are accepted on reinstallation.

`loader-lifecycle.iss` holds the installation checks and exact registration
operations. Compile with `/DSmokeTest` only for automated tests: that build has a
separate uninstall identity, shortcut group and synthetic registry branch under
HKCU/Software/NADABAS/InstallerSmoke. Run `test_loader_installer_lifecycle.ps1`
only against that variant. It exercises the actual installer code without
registering anything in Excel. Do not distribute the smoke variant to users.
The normal package uses the newly signed 0.4.0 session-only XLL.

This is not a transactional production updater: recovery from interrupted
replacement and an Excel process starting during the installation still need
broader acceptance. Process checks at preparation, installation and registration
reduce that race but do not create an atomic Excel-startup interlock.

Local package verification on 2026-10-29 passed 10 checks: installation, both
slot hashes/read-only attributes/absence of zone metadata, valid publisher
signatures on Setup/XLL/Uninstall, A-B-A pointer selection and uninstall cleanup.
The first build did not preserve read-only attributes; explicit Inno attributes
were added and the final signed build passed the checks. These tests used an
isolated workspace destination, did not start Excel, and did not simulate an
Internet download. The delivery and reports are in ignored
`artifacts/loader-installer/0.3.0`. Manual Excel acceptance is still pending.

References: [Inno SignTool](https://jrsoftware.org/ishelp/topic_setup_signtool.htm)
and [signed uninstaller](https://jrsoftware.org/ishelp/topic_setup_signeduninstaller.htm).

Installer 0.3.2 passed 13 checks against the isolated compiled variant on
2026-10-29: unknown-directory refusal, initial registration, unrelated-entry
preservation, active pointer/report preservation on reinstall, exact old payload
registration cleanup, no duplicate loader entry, refusal of modified payload,
own-registration removal and report preservation on uninstall, reinstall with
residual reports/baseline files, and uninstall registration removal. The final
normal variant is signed and verified and is delivered in
`artifacts/loader-installer/0.3.2`. Actual Excel activation remains a user-session
acceptance check, not a result claimed by these isolated registry tests.
