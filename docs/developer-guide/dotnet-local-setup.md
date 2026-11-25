# Local setup for the gradual .NET migration

This guide covers the first read-only diagnostic component. Use the existing
NADABAS repository; the decision and alternatives are recorded in
[ADR 0002](../decisions/0002-incremental-dotnet.qmd).

## Required tools

| Tool | Purpose |
|---|---|
| Git | Review changes and reproduce the source baseline. |
| .NET 10 SDK | Build, run and test C#. The SDK includes the runtime. |
| VS Code with C# support, or a compatible Visual Studio | Optional editor/debugger; CLI builds do not require the IDE. |
| Python 3.13 recommended (project minimum 3.12) | Match the existing Python CI baseline. |
| Poetry 2.2 or later | Install the repository Python dependencies and development tools. |
| Quarto CLI | Render and review the technical documentation. |
| Windows and desktop Excel | Execute VBA and check add-in behaviour. |
| Access/ACE provider matching the process using it | Test the supported Access configuration; inventory alone does not load the driver. |
| SQL Server test instance and the applicable client provider | Test the SQL Server configuration with non-production data. |
| Approved Excel template/release and representative test workbooks | Import, compile and exercise the VBA source. |

The first diagnostic build needs only Git and the .NET SDK. Full migration
acceptance also needs the Windows/Excel/database environment. Install database
providers according to the agreed Office compatibility matrix; do not install
both bitnesses indiscriminately. Confirm Excel bitness in **File > Account >
About Excel**. It is independent of Windows and the diagnostic process bitness.

Install the SDK using the official installer or:

```powershell
winget install --id Microsoft.DotNet.SDK.10 --exact
```

See [Microsoft's Windows installation guide](https://learn.microsoft.com/en-us/dotnet/core/install/windows).
Use your organization's approved installation channel where applicable.
For Python tooling, follow [Poetry installation](https://python-poetry.org/docs/#installation);
for documentation, use [Quarto installation](https://quarto.org/docs/get-started/).

## Automated .NET checks

From the repository root:

```powershell
Set-Location dotnet
dotnet --list-sdks
dotnet restore Nadabas.slnx
dotnet build Nadabas.slnx --configuration Release --no-restore
dotnet test Nadabas.slnx --configuration Release --no-build
dotnet format Nadabas.slnx --verify-no-changes --no-restore
dotnet run --project src/Nadabas.Setup --configuration Release --no-build -- diagnose
Set-Location ..
```

Tests cover filename compatibility, multiple candidates, missing/invalid
directories, preservation of files, argument validation and the JSON contract.
Installation-plan tests also cover fresh installs, unchanged files, replacement
with a required backup, conflicting add-ins and blocked source/target paths.
The .NET workflow runs on Windows and Linux; Linux must report Windows
inventory as unsupported. Passing tests do not establish Excel compatibility.

For restricted build environments, a single-process build can help diagnose
worker-process failures: add `-m:1 --disable-build-servers` to `dotnet build`.
Do not change security policies to make a build pass.

## Existing repository checks

From the repository root, after installing Poetry and Quarto:

```powershell
poetry install
poetry run pre-commit install
poetry run pre-commit run --all-files
poetry run pytest -v
poetry run python tools/scan_vba_security.py vba --approvals tools/vba-security-approvals.json --show-approved
quarto render docs --to html
```

`pre-commit` is installed by `poetry install`; a separate global installation is
unnecessary. Hooks can format files, so review the diff afterwards. See the
[import guide](import-vba.md) and [export guide](export-vba.md) for the approved
template workflow. Enable VBA project object-model access only where required
for development import/export, not as a general end-user prerequisite.

## Build a standalone diagnostic executable

From `dotnet/`:

```powershell
dotnet publish src/Nadabas.Setup -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o ../artifacts/setup/win-x64
../artifacts/setup/win-x64/Nadabas.Setup.exe diagnose
```

Use `win-x86` or `win-arm64` and a distinct output directory only when required
by the deployment profile. This separate process is not loaded into Excel;
its architecture does not have to match Excel merely to exchange files. Any
future in-process COM bridge or database provider has its own bitness needs.
The package is a diagnostic executable, not an installer or signed release.
Bundled runtimes must be republished when security updates are adopted.

See [Microsoft's single-file deployment guide](https://learn.microsoft.com/en-us/dotnet/core/deploying/single-file/overview).

## Manual acceptance checks

For the read-only installation plan, run from `dotnet/` with an actual extracted
release file and the intended Excel add-in directory:

```powershell
dotnet run --project src/Nadabas.Setup -c Release -- plan-install --source "C:\NADABAS-release\NADABAS.xlam" --addin-directory "$env:APPDATA\Microsoft\AddIns"
```

See the [plan contract and limitations](../../dotnet/README.md#installation-planning-read-only).
Check `steps`, `blockers` and `requiredChecks`. A successful plan is still
`review-required`, with `canApply: false`; it does not install or back up files.
In disposable test folders, compare plans for absent, identical and differing
targets and a second versioned add-in. Confirm no files or directories change.
Also test an inaccessible folder and a junction/symbolic link; both must block.
The source must come from the approved release channel before future installation.

Record the source commit, Windows/Office version and bitness, provider versions,
database engine, test case, expected result and actual result.

1. Run diagnostics on a clean test profile, an existing NADABAS profile and a
   custom add-in directory. Compare observations with Excel's add-in list.
2. Use temporary directories containing zero, one and multiple XLAM candidates.
   Confirm the report is accurate and files are unchanged. Repeat with an
   inaccessible directory and confirm an incomplete report, not a clean bill
   of health. Check paths containing spaces and non-ASCII characters.
3. Run the published executable on a test machine without the .NET runtime.
   Verify JSON output and exit codes. Review path information before sharing.
4. In a copy of the approved Excel template, import the reviewed VBA and run
   **Debug > Compile VBAProject**. Verify startup, ribbon and shutdown.
5. For each supported Access/SQL Server profile, exercise database opening,
   workbook registration/reservation, load/save, classifications and a small
   batch. Compare values, metadata and reservation state with the baseline.
6. Before moving writes into .NET, add interrupted-write, concurrent-user,
   transaction and restore/reconciliation scenarios with expected outcomes.

Future installation acceptance additionally requires clean install, upgrade,
Excel-open/file-lock handling, failed copy, repair, uninstall and rollback.
Signing certificates and publisher trust are release requirements, not needed
to compile this diagnostic prototype. Do not treat CI as evidence that those
release scenarios passed.
