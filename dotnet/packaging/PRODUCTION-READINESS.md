# Release acceptance gates

0.8 engine candidate: signed out-of-process package application is implemented
and undergoing isolated integration tests. Whole-package Excel UI handoff and
transactional/interruption recovery remain incomplete. Contextual Exchange
visibility is retained. See PACKAGE-UPDATER-0.8.md for exact scope.


Current 1.0 scope is recorded in [V1-RELEASE-PLAN.md](V1-RELEASE-PLAN.md).
It includes unattended updating of the updater and loader, not only XLAM.
The signed 0.6.0 test package adds Excel notification and VBA handoff;
the maintainer reported its tests successful on 2026-11-03. The historical limitations below
are not the intended final 1.0 behavior.


The maintainer reported that installation 0.3.2 worked on 2026-11-02. That is
user-profile evidence, not approval for organization-wide deployment.

The maintainer subsequently reported the 0.4.0 test guide completed. Record this
as user-reported upgrade/rollback acceptance; it does not imply clean-machine,
database or production-download validation. GitHub Releases with a signed feed
was selected for distribution, with automatic checks/download preparation and
an in-product restart notification. Users should not visit GitHub manually.

The 0.5.0 milestone implements the pending-activation backend. Queueing does not
change active.json; the next loader startup verifies and atomically applies the
queued approved slot. Failed validation retains the old selected release and
reports rejection. Cancellation and explicit activation clear pending requests.
An interrupted pointer switch can resume idempotently. This only handles releases
already admitted by the compiled allowlist and staged locally; it does not yet
implement an unattended downloader, release admission, UI notification or policy.

Next integration: use the signed GitHub feed to discover a newer approved package,
verify before staging, obey IT disable/pinning policy, expose readiness in English,
and apply without killing Excel. Changes to the loader itself still require the
installer/repair path and Excel closure. Never represent a failed download or
an unverified package as ready to restart. Do not publish the simulated VBA release.

The next internal package, 0.4.0, tests a deliberately simulated VBA 6.02.006.
Only the runtime version constant and an About test label change relative to
the signed 6.02.005 baseline. It must not be published as a normal NADABAS release.
No local business-logic changes are implicitly included.

## Implemented for the next acceptance round

- English installation guide, messages and Start Menu shortcut names.
- Prerequisite checks for installed 64-bit Excel and registered .NET 10 Desktop
  Runtime x64. Missing prerequisites are explained before installation changes.
- Compiled release allowlist binding each slot to an expected version and hash.
  The signed loader verifies file content and the VBA-reported runtime version.
  A mutable adjacent JSON manifest cannot authorize an arbitrary new payload.
- Separate immutable slots with atomic activation and retention for rollback.
- Automatic loader registration, exact cleanup, known-directory repair and
  preservation of unrelated startup entries and reports.
- Authenticode verification of the supplied loader and VBA test fixture during
  packaging, with signed Setup and Uninstall.

## 0.7.0 acceptance pending

Restart and update now uses a separate helper and normal Excel shutdown. Automated
wait/cancellation tests are not evidence that Excel Quit/reopen succeeds under
all add-ins and save prompts. The user must test the included guide before this
is accepted. Workbooks are not restored automatically; a helper timeout retains
normal manual restart as a fallback. The version 008 date is stamped before signing.

## Required before external production distribution

| Gate | Evidence required |
| --- | --- |
| Actual upgrade and rollback | About and session reports show 6.02.005 and simulated 6.02.006 correctly across normal and concurrent Excel processes. |
| Clean machine | Supported Windows/Office matrix, first install without SDK, missing runtime handling, standard user and IT-managed deployment. |
| Download and publisher trust | Install a genuinely browser-downloaded signed package under the organization's Office policy; verify XLL and VBA signatures and prompts. |
| Failure recovery | Interrupted extraction, full disk, access denied, registry failure, process-start race and recovery to a usable previous release. Current package is not a transactional updater. |
| Data regression | Load, save, batch, workbook registration and representative test databases; code rollback is not database rollback. |
| Update channel | Production HTTPS feed, pinned signing trust, key rotation, opt-out/pinning and ownership when IT manages updates. The NetSparkle probe is not yet integrated. |
| Release governance | Reviewed source commit, approved artifact hashes, genuine version number, VBA signing verification and reproducible release record. |
| Support and licensing | Installer-tool licensing, diagnostics policy, user documentation, uninstall/repair instructions and retention policy. |

The test installer may select its packaged release again on repair. Production
repair, rollback and downgrade semantics must be specified separately before
unattended updates are enabled. Do not silently switch between these behaviors.

Prerequisite detection currently covers the registered machine-wide .NET runtime.
Custom or portable runtime layouts are not claimed as supported. The package
does not automatically download prerequisites or weaken Office security policy.

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
