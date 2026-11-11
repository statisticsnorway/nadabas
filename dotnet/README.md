# NADABAS Setup & Diagnostics

First incremental .NET component, built with C# and .NET 10 LTS. This is a
diagnostic/planning prototype with isolated file-installation rehearsals,
not a production installer or a replacement for VBA.
No runtime NuGet dependencies are required; the test project uses xUnit.

That dependency statement applies to `Nadabas.Setup`. The separate
`Nadabas.UpdateProbe` experiment uses NetSparkle 3.1.0. See the
[installation prototype](packaging/README.md) for its test-only scope and
reproducible EXE install/upgrade/uninstall checks.

From this directory:

```powershell
dotnet restore Nadabas.slnx
dotnet build Nadabas.slnx -c Release --no-restore
dotnet test Nadabas.slnx -c Release --no-build
dotnet run --project src/Nadabas.Setup -c Release --no-build -- diagnose
```

The SDK policy in `global.json` permits stable .NET 10 feature-band updates.
Run commands from `dotnet/` so this policy is applied. End users of a future
self-contained release will not need the SDK.

See [local setup and verification](../docs/developer-guide/dotnet-local-setup.md)
for prerequisites, publishing, manual checks and the remaining migration work.

## Published release baseline

The installation migration uses the maintainer-supplied **6.02.005** artifacts:
`NADABAS.6.02.005.xlsm` and `NADABAS.XLAM`. Their measured sizes and SHA-256
hashes are recorded in [releases/6.02.005.json](releases/6.02.005.json).
See the [baseline evidence and update procedure](../docs/operations/release-6.02.005.qmd).
Unreleased local VBA changes remain development work, separate from this baseline.
`rehearse-install` checks its source against the embedded 6.02.005 add-in hash;
`plan-install` does not yet enforce the manifest or verify Office signatures.
Future releases get a new manifest rather than
overwriting the 6.02.005 hashes.

## Diagnostic contract, version 1

```powershell
dotnet run --project src/Nadabas.Setup -c Release --no-build -- diagnose --addin-directory "C:\My NADABAS AddIns"
```

The optional path must be absolute. The default is the conventional current-user
`%APPDATA%\Microsoft\AddIns` directory, not a confirmed Excel
`Application.UserLibraryPath`. Custom locations must be supplied explicitly.
Only the immediate directory is enumerated. XLAM files are identified by the
existing installer filename convention; their contents are not opened.

The built executable writes JSON to stdout, with `schemaVersion`, `toolVersion`,
`operatingSystem`, `processArchitecture`, `complete` and `findings`. Each finding
has `id`, `status`, `evidence` and an optional `value`. Registry view names and
the tool's process architecture must not be interpreted as Excel bitness.

| Status | Meaning |
|---|---|
| `observed` | File candidate or registration found, not a functional verification. |
| `not-found` | Not found in this particular location or registry view. |
| `warning` | Multiple candidate add-ins need review. |
| `not-tested` | Requires a separate functional check. |
| `error` | An inventory operation failed, for example access was denied. |
| `unsupported` | Windows registry inventory cannot run on this OS. |

Exit codes: `0` means inventory completed, **not installation readiness**;
`1` means inventory was incomplete; `2` means invalid arguments (stderr, no
JSON report). Missing registrations are observations, not fatal errors.
`complete` is false for `error` or `unsupported`. Reports include local paths;
review them before sharing and store saved reports under ignored `artifacts/`.

On Windows, the tool reads only Excel App Paths and ACE 12.0/16.0 ProgID
registrations in applicable registry views. It does not read NADABAS database
settings or credentials, start Excel, contact a server, change trust settings,
verify signatures, register add-ins, install drivers, or mutate files/registry.
This describes `diagnose`; the separate rehearsal command below writes test files.
SQL connectivity and installed NADABAS version are not determined yet.

## Installation planning (read-only)

From `dotnet/`, supply the exact release XLAM and intended Excel add-in directory:

```powershell
dotnet run --project src/Nadabas.Setup -c Release -- plan-install --source "C:\NADABAS-release\NADABAS.xlam" --addin-directory "$env:APPDATA\Microsoft\AddIns"
```

Replace the source example with the actual extracted release file. The source
must have a NADABAS XLAM filename; both paths must be absolute. The directory
is explicit because the conventional location is not proof of Excel's configured
location. The command hashes files without opening Excel or changing any files.

The JSON plan has `kind: installation-plan`, `schemaVersion: 1`, source and
target paths, SHA-256/size snapshots, proposed `steps`, `blockers` and
`requiredChecks`. Its contract is distinct from `diagnose`. A successful plan
has `status: review-required` and **always `canApply: false`**. There is no apply
command in this increment. Exit code 0 means a reviewable plan, 1 a blocked
plan, and 2 invalid command syntax or paths.

| Situation | Proposed result |
|---|---|
| Target absent | Install the source as `NADABAS.XLAM`; no directory is created now. |
| Target has different bytes | Back up the existing file, then replace it. |
| Target has identical bytes | Keep it; activation and compatibility still need checking. |
| Other versioned NADABAS add-ins present | Block; inspect registrations in Excel before deciding how to handle them. |
| Missing/empty source, unreadable paths, links/junctions or invalid target | Block without proposed installation steps. |

Backup destinations are proposals below `.nadabas-backups/<old-SHA256>/` in the
target directory. They are not created or checked for availability by this
planner. A future installer must validate their paths and permissions, preserve
or verify any existing backup, and verify the saved bytes before replacement.
It must recheck hashes, links and locks at execution time: a saved plan can
become stale immediately. This prototype conservatively rejects reparse points,
which can include cloud-managed locations; use a normal local release directory.

Hashes compare content, **not versions or trust**. A differing file can be older
or newer. Package contents, VBA signatures, publisher trust, Excel registration,
file locks and write permissions are explicitly pending checks. Even arbitrary
nonempty bytes with the right filename can produce a review-required plan.
No file is represented as a verified or approved release on that basis.

## Stage 3: file installation and rollback rehearsal

The maintainer reported successful manual signature review of 6.02.005 before
this stage. Automated Office signature validation is still separate work.

From `dotnet/` in PowerShell:

```powershell
$testOutput = [IO.Path]::GetFullPath((Join-Path $PWD '../artifacts/install-rehearsal'))
dotnet run --project src/Nadabas.Setup -c Release -- rehearse-install --source "C:\NADABAS\NADABAS.XLAM" --output "$testOutput"
```

This command validates the source against the embedded reference add-in hash
before creating test files. It creates a new `nadabas-rehearsal-<random-id>`
directory for each of five scenarios, never using an existing directory as
the installation target. Sources and the active Excel add-in are not modified.
There is no arbitrary target-directory apply command.

| Scenario | Expected status and final file |
|---|---|
| `fresh-install` | `installed`, reference bytes present |
| `identical` | `unchanged`, reference bytes preserved |
| `replace-and-restore` | `rolled-back`, synthetic previous bytes restored |
| `failure-after-replace` | `rolled-back`, injected I/O failure triggers automatic recovery |
| `interrupted-and-recovered` | `rolled-back`, a new transaction object recovers from the persisted journal |

The previous file in replacement tests is explicitly synthetic, not another
release. A successful JSON report has `passed: true`, five case results with
`passed: true`, and exit code 0. Operation failure is exit 1; malformed command
syntax is exit 2. Test folders and verified backups remain for inspection.

Each operation rechecks source and target hashes, stages and verifies the new
file, copies and verifies the old file before replacement, writes a journal,
then renames the staged file into place and verifies the final hash. Existing
backups with different contents are rejected rather than overwritten. Read-only
or locked targets are not forcibly unlocked. Failed operations before replacement
preserve the target; failures after preparation attempt rollback. Unexpected
third-party modifications or damaged backups require manual recovery and are
preserved. A fresh-install rollback moves the new file aside rather than deleting it.

To repeat recovery on a generated test workspace, use its path from the report:

```powershell
dotnet run --project src/Nadabas.Setup -c Release -- recover-test --workspace "C:\path\nadabas-rehearsal-<id>"
```

Only generated workspace names with the expected marker are accepted. Recovery
reads the journal, verifies hashes and never takes file paths from the journal.
Successful recovery returns `rolled-back`; repeating it is safe if the files
are unchanged. Use a new rehearsal for a new transaction. Do not edit journals,
copy workspace markers into other directories, or register the test files in Excel.

Tests cover stale plans, locked/read-only targets, failure checkpoints, corrupt
backups, changed staging/target bytes, persisted recovery and invalid journals.
The interruption is simulated; this does not establish power-loss durability,
Windows reboot behaviour, ACL preservation, or recovery from actual process
termination. The prototype assumes exclusive use of its disposable directories;
path rechecks do not defend against a hostile process racing filesystem changes.
Production integration and Office activation remain stage 4 work. No Excel
instance or database is opened during these tests.

## Stage 4: Excel activation check (Windows)

The initial `Workbooks.Open` probe could run VBA without displaying the Ribbon.
Its old `passed` reports are not evidence of add-in activation. The corrected
test uses `AddIns.Add` and `Installed = true` and checks the exact registered path.
It refuses to start while Excel is running or another NADABAS add-in is active.

From `dotnet/` after saving your work and closing Excel:

```powershell
$testOutput = [IO.Path]::GetFullPath((Join-Path $PWD '../artifacts/excel-integration/activation'))
dotnet run --project src/Nadabas.Setup -c Release -- excel-startup-test --source "C:\NADABAS\NADABAS.XLAM" --output "$testOutput" --interactive
```

The copied file is named `NADABAS.TEST.<random-id>.xlam`. Excel's window caption
identifies the test session. The terminal prints the full active test path.
Check that the NADABAS Ribbon tab appears and About shows 6.02.005. In Excel's
add-in list, identify the test by its **path**, not its title (the signed workbook
title may still be Nadabas). Do not activate the normal AppData copy.

The test follows existing Office macro policy. It does not enable all macros,
change trusted locations, re-sign the file, or open any database. Startup uses
the current Windows profile and may read existing NADABAS settings; this is not
a virtual machine. The scheduled update check is cancelled after startup.

Close dialogs, then press Enter in the command's terminal (or tell the assistant
you are finished when it owns the test terminal). The helper deactivates its
test registration and removes only exact references to its own unique path
from Excel's current-user Add-in Manager/OPEN entries after Excel closes.
It does not alter the original NADABAS registration. Avoid closing Excel with X
mid-test; unexpected closure requires checking cleanup errors in the report.

Schema version 2 reports `addinActivated`, `registrationCleaned` and
`ribbonVerification`. `activation-verified-ribbon-pending` and exit code 0 mean
the programmatic checks succeeded; **the Ribbon still requires user observation**.
An `excel-startup-report.json` is saved beside the copy on normal completion.
Database loading/saving, batch work and persistent installation remain separate
acceptance steps. The practice package is copied under
`artifacts/excel-integration/practice-copy`; no database has been opened by this test.

## Next increments

### Isolated update rehearsal

Run from `dotnet/`:

```powershell
$testOutput = [IO.Path]::GetFullPath((Join-Path $PWD '../artifacts/update-rehearsal'))
dotnet run --project src/Nadabas.Setup -c Release -- rehearse-update --source "C:\NADABAS\NADABAS.XLAM" --output "$testOutput"
```

This tests the update coordinator with the approved 6.02.005 incoming file and
explicitly synthetic previous bytes, not two released versions. Each scenario
gets a new disposable workspace. The 11 reported checks cover replacement and
backup, deferral and retry after Excel closes, a second Excel-state check before
replacement, IT enable/disable precedence, version pins, user opt-out, identical
content and rollback after an injected replacement failure. Exit 0 requires all
checks and an unchanged source hash. Reports include hashes and retained paths.

Excel state and IT policy are injected test inputs: no Excel instance is opened,
no Windows policy is installed, and no download or automatic scheduling happens.
The coordinator is test-only. A process check is not an interlock against Excel
starting concurrently. Production work still includes this race, actual process
detection, authenticated packages, policy delivery, version ordering, component
compatibility and regression testing between real releases. These tests do not
establish that database operations work after an upgrade.

### Live Excel process rehearsal (Windows)

Save your work and close Excel, then run from `dotnet/`:

```powershell
$testOutput = [IO.Path]::GetFullPath((Join-Path $PWD '../artifacts/live-update-rehearsal'))
dotnet run --project src/Nadabas.Setup -c Release -- rehearse-update-live --source "C:\NADABAS\NADABAS.XLAM" --output "$testOutput"
```

This creates its own Excel COM instance without opening workbooks or activating
the test add-in. It checks actual Windows EXCEL processes, verifies that the
isolated update is deferred without a transaction journal or backup, closes its
own empty instance, retries, verifies the installed/backup hashes, and restores
the synthetic previous bytes. Existing Excel processes cause refusal before
test preparation. Unexpected workbooks are left open for user inspection.
Detection errors stop the update; they are never interpreted as an empty list.

Exit 0 requires all three checks, cleanup and an unchanged release source. JSON
output identifies real process detection and the still-simulated policy inputs.
This runs in the current Windows profile; it is not a clean virtual machine.
The process checks still cannot prevent Excel starting just after the final
check. That interlock, actual IT policy, scheduling and signed package delivery
remain production work. No installed add-in or database is modified.

### Remaining implementation

The update coordinator now rechecks Excel immediately before the rename, after
staging and backup. A late start or a failed process query cancels replacement
and preserves the previous file. Rollback results retain the triggering error.
Windows tests also cover a reader opening the target after the final query,
and a reader holding the new file during failed-update recovery. Recovery never
forces that reader closed: it reports `recovery-required`, preserves the backup,
and can be retried after the reader releases the file.

These tests use real Windows file handles, not Excel loading an XLAM during the
rename. A trial that held read-denying handles across `File.Move` also blocked
the updater's own rename on this machine and was not retained. Windows sharing
constraints are documented in [CreateFile](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createfilea).
The last process-check/rename race remains open. A coordinated loader or another
deployment design needs evaluation before automatic production updates; this
increment does not claim an Excel-startup interlock or a seamless loading experience.

1. Record a supported Office/Windows/driver matrix and verify diagnostics there.
2. Extend publisher verification and harden the tested file transaction for
   production paths, process interruption and installation metadata.
3. Replace the XLSM installation workflow after clean-install, upgrade,
   failure and rollback tests pass with signed release artifacts.
4. Add a bounded VBA-to-process adapter and move version checking behind a
   versioned request/response contract. Define timeouts and failure behaviour.
5. Extract pure validation rules; migrate complete database operations later.

The production VBA path remains unchanged in this increment.
