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
  diagnostics candidate passed 137 .NET tests and 129 Python tests. The 0.9.5
  components were signed; the subsequent 1.0 engineering build deliberately
  remains unsigned. Full multilingual Details UI acceptance remains open.

Historical test packages contain simulated/beta payloads and must stay out of
the stable channel. The 1.0 production profile ships the selected 6.02.009 and
baseline slots only, with a NADABAS display name and compatible internal identity.
Its unsigned engineering output is not distributable. Existing signed artifacts do not include later
source changes. See [AUTOMATIC-UPDATES.md](AUTOMATIC-UPDATES.md) for runtime
evidence and [RELEASE-SIGNING.md](RELEASE-SIGNING.md) for the release procedure.

## Required before external production distribution

| Gate | Remaining evidence or decision |
| --- | --- |
| Production package | VBA 6.02.009 selected. Compact production profile and compatible identity implemented; complete signed 1.0 and ordinary desktop acceptance. |
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
a copy of the approved 6.02.005 XLAM, an explicitly selected signed smoke installer and synthetic
current-user startup entries under Software/NADABAS/InstallerSmoke. The backup
and registry snapshot are outside the installation directory. The rehearsal
performs the switch and rollback around Setup. The shared Setup engine implements this handoff. On 2026-11-12 the signed
0.9.5 smoke package passed 14 Setup migration checks, including preserved
registration, refused ambiguous inputs and exact restoration. Ordinary desktop
legacy-only acceptance remains separate.

For Excel acceptance, use a disposable Windows profile or VM and a non-production
database. Start Excel normally with the old add-in first. After migration, start
Excel normally again and check the loaded file path and version, one NADABAS tab,
and another enabled add-in. Restore the original registration, restart Excel and
confirm the old file path and version. Capture the session evidence for both starts.

## Token-free 1.0 preparation, 2026-11-12

See [V1-PREPARATION.md](V1-PREPARATION.md) for the build/signing boundary and
[SSH-SETUP.md](SSH-SETUP.md) for the next maintainer setup session.

- Unsigned 1.0 build: 15 isolated installer checks passed, including refusal
  to install the unsigned ordinary-profile artifact and all four language codes.
- Existing signed 0.9.4/0.9.5 fixtures: 29 package checks passed, including
  a real failed installation, rejection of the wrong recovery version, recovery
  to 0.9.4, then upgrade to 0.9.5. This is not signed 1.0 acceptance.
- Existing signed 0.9.5 fixture: 14 legacy Setup/Uninstall checks passed.
- Repository pre-commit checks and 129 Python tests passed. No VBA or .NET
  runtime implementation changed in this preparation; the previous 137 .NET
  tests remain applicable, and the 1.0 components built successfully.
- No new token signature, public release or normal-profile installation was made.

Actual GitHub delivery, clean-machine/Office policy, representative data
regression, key backup/recovery and installer licensing are still open. Do not
reduce the remaining work to signing alone.
