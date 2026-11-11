# Release acceptance gates

Reviewed against source and recorded acceptance on 2026-11-11.
The target scope is [V1-RELEASE-PLAN.md](V1-RELEASE-PLAN.md). This checklist
is not production approval. Historical prototype numbers do not establish
acceptance of the current build.

## Implemented and evidenced

- Per-user installation, loader registration, prerequisite checks, language
  selection and immutable XLAM slots with verified activation and rollback.
- Signed-feed discovery and package download, NetSparkle strict signature
  verification, publisher checks, bounded HTTPS transport and IT update policy.
- Excel ready notification and external whole-package restart/install helper.
  The ordinary desktop test installed package 0.9.3 and reopened Excel with VBA
  6.02.006. Its previous Windows registration was 0.8.2, while the loaded
  component was 0.9.2. This was a locally staged package, not GitHub delivery.
- SSH source verification, Ed25519 key overlap/revocation and encrypted recovery
  tooling. Automated recovery uses temporary keys; organizational custody and
  a real replacement-workstation recovery drill remain open.
- Local TLS coverage exercises signed feed discovery and the GitHub asset
  redirect shape. It does not establish access to published GitHub assets or
  successful installation of those assets.
- Update diagnostics and translated Details UI are implemented. The 2026-11-11
  run passed 133 .NET tests and 127 Python tests and built the Excel bootstrap.
  This UI change is unsigned and still needs visual Excel acceptance.

The current packages retain test identities and simulated/beta payloads. Keep
these out of the stable channel. Existing signed artifacts do not include later
source changes. See [AUTOMATIC-UPDATES.md](AUTOMATIC-UPDATES.md) for runtime
evidence and [RELEASE-SIGNING.md](RELEASE-SIGNING.md) for the release procedure.

## Required before external production distribution

| Gate | Remaining evidence or decision |
| --- | --- |
| Production package | Select the approved VBA payload, genuine versions and final installation identity; exclude simulated fixtures. Define migration from both legacy XLAM and test installations. |
| Actual distribution | Exercise the installed updater against an approved published GitHub feed: discover, download, verify, notify, restart, install and relaunch. Publication requires explicit approval. |
| Clean machine | Agree the Windows/Office support matrix. Test standard-user installation without the SDK, prerequisite failures and IT-managed deployment. |
| Browser and publisher trust | Test a browser-downloaded final signed installer under the intended Office policy, including XLL/VBA trust and prompts. |
| Restart and recovery | Check unsaved work, concurrent Excel processes, Later/cancel, interrupted extraction, full disk, denied access, registry failure and process-start races. Prove explicit recovery to a usable previous package. Whole-package application is not transactional. |
| Legacy migration | Rehearse a legacy-only ordinary Windows profile, backup, registration handoff and exact restoration. Keep user workbooks and databases outside replacement. |
| Data regression | Load, save, batch and workbook registration on representative test databases. Code rollback does not roll back data. |
| Languages and accessibility | Inspect all four languages in Excel, first-use selection, keyboard operation and larger display scales, including the new Details window. |
| Release governance | Approved SSH signer list, reviewed source, final signed artifact hashes, production feed-key custody, replacement-workstation recovery and a rehearsed rotation route. |
| Support and licensing | Resolve installer-tool licensing, support ownership, repair/downgrade semantics and user instructions. |

For desktop acceptance, verify both physical cache paths and ordinary Windows
registration. Packaged Codex processes have previously observed different views.
Preserve package verification while diagnosing that discrepancy.

Prerequisite detection covers the registered machine-wide .NET Desktop Runtime.
Custom or portable runtime layouts are not claimed as supported. Setup does not
automatically download prerequisites or change Office security policy.

## Migration from the VBA installer

The former installer copies NADABAS.XLAM into Excel's UserLibraryPath. Run
Nadabas.Setup plan-legacy-migration in the user's profile to collect candidate
files, their hashes, registration evidence and startup-folder findings. This
read-only inventory is the first migration step. An Add-in Manager entry alone
does not establish that the add-in is active.

Before enabling automatic migration, rehearse a legacy-only profile. With Excel
closed, verify a backup of the identified file and the original registry values.
Install the loader in the managed per-user directory. Switch only the confirmed
legacy startup entry and verify that one NADABAS ribbon appears on normal Excel
startup. Redirect custom, shared and ambiguous paths to review. Test a reversible
move for any auto-loaded XLSTART copy and restoration of the exact previous
registration if activation fails. Keep user databases and workbooks in place.

Run `tools/test_legacy_migration.ps1 -SmokeInstaller <signed smoke installer>`
to rehearse the registry handoff on Windows with Excel closed. The harness uses
a copy of the approved 6.02.005 XLAM, the 0.8.1 smoke installer and synthetic
current-user startup entries under Software/NADABAS/InstallerSmoke. The backup
and registry snapshot are outside the installation directory. The rehearsal
performs the switch and rollback around Setup. Automatic legacy handoff in the
user-facing Setup remains a separate integration step.

For Excel acceptance, use a disposable Windows profile or VM and a non-production
database. Start Excel normally with the old add-in first. After migration, start
Excel normally again and check the loaded file path and version, one NADABAS tab,
and another enabled add-in. Restore the original registration, restart Excel and
confirm the old file path and version. Capture the session evidence for both starts.
