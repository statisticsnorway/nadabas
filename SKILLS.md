# NADABAS workflow handbook

Maintainer review draft, revised 2026-11-11.
[AGENTS.md](AGENTS.md) directs agents here. Read the workflow relevant to the task.
This repository handbook is read through that instruction. Independently
discoverable Codex skills use a skill directory with a `SKILL.md` entrypoint.

## Task routing

| Change | Workflow |
|---|---|
| Existing behaviour or migration | Incremental VBA and .NET work |
| Database writes, permissions, batch or schema | Data and authorization |
| Forms, ribbon, language or user workflow | Languages and Excel UI, plus DESIGN.md |
| Python/.NET implementation or build checks | Build and automated verification |
| Installation, updates or release trust | Relevant platform workflow below |
| Documentation and onboarding | Documentation and technical claims |

## Incremental VBA and .NET work

Use when changing an existing function or moving responsibility to .NET.

1. Identify the VBA entry point, its callers, side effects, and persisted state.
2. Define the behaviour to preserve and the small responsibility being moved.
   Inspect existing bridges and shared logic before introducing another interface.
3. Exercise normal, missing-input, and failure paths. Include an existing workbook
   or older persisted state where compatibility matters.
4. Run relevant automated checks, then validate Excel-specific behaviour in a
   test copy. Record database changes and recovery expectations explicitly.

Start with [the migration decision](docs/decisions/0002-incremental-dotnet.qmd),
[VBA import](docs/developer-guide/import-vba.md), and
[VBA export](docs/developer-guide/export-vba.md).

## Data and authorization

Use for database access, identity, privileges, load/save, batch and schema work.
Read [database access](docs/vba/database-access.qmd),
[workbooks and data](docs/vba/workbooks-and-data.qmd), and
[administration](docs/vba/domain-and-administration.qmd).

1. Identify the actor, active database/workbook, action, permission source,
   transaction boundary and consequences of partial completion.
2. Trace privileged operations through direct calls as well as UI entry points.
   Check stale permission state after switching database, workbook or identity.
3. Verify test-data identity and recoverability before writes or schema changes.
4. Check relevant SQL parameterization, reservation/conflict behaviour, locks,
   interrupted writes and retry semantics. Determine the previous result before
   replaying an operation that may have committed.
5. Verify the actual backend. Access file permissions and SQL Server
   authorization can provide different guarantees.

The legacy empty-administrator-list rule is recorded in DESIGN.md.
Changes to it require an explicit compatibility and security decision, covering
bootstrap, existing databases, failed initialization and administrator recovery.

## Build and automated verification

Use when changing source, tests, packaging, or the release tools.

Select tests by changed behaviour during development. Before a PR, complete the
checks required by CONTRIBUTING.md and the relevant GitHub workflows.
Read script parameters before executing acceptance scripts. Some scripts install
software, change registration, or open Excel.

Python checks run from the repository root using the configured environment:

```powershell
poetry run pytest -v
poetry run python tools/scan_vba_security.py vba --approvals tools/vba-security-approvals.json --show-approved
```

Select the actual affected Python test files for the task.
Use focused test selection during development and the full suite before submission
as required by CONTRIBUTING.md.
From `dotnet/`, the standard .NET checks are:

```powershell
dotnet restore Nadabas.slnx --locked-mode
dotnet build Nadabas.slnx -c Release --no-restore
dotnet test Nadabas.slnx -c Release --no-build
```

The Excel bootstrap and release signing tool have separate project files.
Build these explicitly when changing their code or dependencies. Run the
formatting checks specified by [.NET CI](.github/workflows/dotnet.yml).

If sandbox networking, test-host IPC, or temporary directory permissions fail,
record the environmental failure and use the supported escalation mechanism or
a suitable workspace directory. Keep expected assertions and security checks.
Use a short local build path if OneDrive or path length prevents packaging.
Verify artifact hashes when copying completed outputs back to the workspace.
An alternate launch path must respect the same authorization and security controls.

For performance work, capture a baseline and representative workload. Measure
the changed path and check correctness alongside timing and resource usage.

## Installer, legacy migration, and restart acceptance

Use when changing installation, registration, version selection, or restart.

Read [the loader installer procedure](dotnet/packaging/LOADER-INSTALLER.md) and
[automatic updates](dotnet/packaging/AUTOMATIC-UPDATES.md).
Reuse the relevant scripts:

- `tools/build_loader_test_installer.ps1`
- `tools/test_loader_installer_lifecycle.ps1`
- `tools/test_legacy_migration.ps1`
- `tools/test_package_update.ps1`
- `tools/test_platform_restart.ps1`

Identify whether the run uses an isolated installation or the user's normal
profile. Check Excel sessions, the current installation identity, expected
source hashes, target directories, and any pending update before mutation.
Preserve logs and backups needed to diagnose an incomplete operation.

For a normal-profile upgrade, verify both the physical cache path and the
Windows registration as seen by the ordinary desktop session. If Codex
virtualization gives a different view, establish a normal-user launch path.
Use narrowly scoped, temporary helpers when required and remove their temporary
registration after completion.

Exercise discovery, verification, the ready notification, restart acknowledgement,
normal Excel exit, installation, relaunch, and loading the intended XLAM.
Check the installed package version and the actual session report. An About
dialog proves only the version it displays.

Include unsaved-work handling, another Excel process, Later, cancellation,
failed installation, and explicit recovery where affected by the change.
Verify that legacy registration leaves the intended loader active and that user
workbooks and database locations remain usable.

## Signing keys and release preparation

Use when preparing artifacts, changing trust, rotating keys, or testing recovery.

Follow [release signing and recovery](dotnet/packaging/RELEASE-SIGNING.md).
The operational tools are:

- `tools/verify_release_source.py` for SSH commit verification.
- `tools/prepare_verified_feed.py` for reviewed release-feed preparation.
- `dotnet/tools/Nadabas.FeedSigning` for feed signing and key backup/recovery.

Verify the release source against the independently approved allowed_signers
file. Build from that reviewed source and record the association between source,
artifact hashes, signatures, and versions.

Use ephemeral keys for automated tests. Before touching a real key, identify its
role, custodian, and intended action. A planned rotation needs a bridge release
and a usable feed for clients still trusting the old key. Test rejection of
revoked and unknown keys as well as successful verification.

When the GlobalSign token is unavailable, build unsigned components, run the
independent checks, and prepare an explicit signing handoff. Sign final bytes,
verify their signatures, then calculate the final distributable hashes.

Treat public publication as a separate, authorized step. Check that the package
contains the approved release payload and that simulated VBA fixtures cannot
reach the stable channel.

## Automatic download tests

Use when changing feed parsing, HTTP transport, trust rules, or update scheduling.

Read `SignedFeed.cs`, `AutomaticUpdates.cs`, and `UpdatePolicy.cs` under
`dotnet/src/Nadabas.PackageUpdate/`. Start with the signed-feed, release-key,
and transport tests under `dotnet/tests/Nadabas.Setup.Tests/`.

Exercise valid and altered feeds/packages, unknown and revoked keys, version
selection, IT pins, cancellation, download bounds, and allowed redirects as
relevant. Keep local test-certificate handling confined to the test transport.

Record whether evidence comes from an HTTP handler, actual local TLS traffic,
or a published GitHub asset. Full acceptance follows the installed updater
through download and staging to the Excel notification and restart.

## Languages and Excel UI

Use when changing first-run language choice, shared dialogs, ribbon callbacks,
or optional Exchange functionality.

Use [DESIGN.md](DESIGN.md) to identify the user, their task and permission scope.
For substantial new flows, present the simplest viable interaction for review.
Implement ordinary small UI improvements within the requested scope directly.

Read the current `vba/resources/office-ui.csv`, language-selection code, and
`dotnet/src/Nadabas.ExcelBootstrap/update-text.json`.
Reuse `tests/test_update_languages.py`, `tests/test_exchange_visibility.py`,
and `tools/test_demo_languages.py` where applicable.

Check first installation and existing-user preferences in English, French,
Portuguese, and Bahasa Indonesia. Inspect actual dialog and ribbon layout with
longer translations. Test optional Exchange controls with and without usable
Exchange links, following the current accepted behaviour.

Report which languages and interactions were actually exercised. Keep pending
visual checks visible even when resource and automated tests pass.

Include keyboard navigation, larger text, understandable errors, and context
changes where affected. Optional tools need a discoverable first-use path.

## Documentation and technical claims

Identify the reader and what they need to decide or do. Use actual source names
and link to the authoritative procedure rather than duplicating it.
Distinguish intended design, source observations and verified runtime results.
Describe production readiness and security guarantees only to the extent supported
by the evidence.

For Quarto changes, validate through the documentation workflow and inspect
rendering where layout changed. For Markdown guidance, check links, referenced
paths and consistency. Review examples and reports for secrets, identifying
data and dangerous commands before sharing them.
