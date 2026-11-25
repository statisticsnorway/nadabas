# NADABAS installation and updates 1.0

Agreed scope, 2026-11-03: the update mechanism itself must be updateable.
Version 1.0 describes the delivery platform, not the VBA application version.
The first supported target is Windows with 64-bit Excel.

## One delivery path for all components

A signed release package contains the compatible XLAM, Excel loader, updater,
and version-management tools. GitHub Releases hosts the packages and a signed
feed; users do not need to visit GitHub or rerun Setup manually for each release.
NetSparkle discovery/download verification is the existing integration starting
point. Its loopback probe is not a production updater.

An updater outside Excel prepares an authenticated package in a separate cache.
It honors administrator policy before discovery, download and installation.
Only a complete verified package can produce the Excel restart notification.
The user saves work and closes Excel; the helper launches the signed installer
from the cache after the old updater exits. The package can replace the updater
as well as the XLL. The installer must not execute from the directory it replaces.
No force-closing Excel and no automatic reopening of user workbooks.

The installer must refuse or defer if Excel starts again. File locks are the last
barrier, not an excuse to continue with a mixed release. Interrupted upgrades
need recoverable state and explicit post-install verification. The current Inno
prototype has guards but does not yet provide full transaction recovery.

Existing per-version XLAM slots remain useful for application rollback. A failed
platform update also needs recovery of compatible loader/updater binaries and
registration. Database contents are not rolled back by either mechanism.

## Milestones and acceptance

1. 0.6: signed Excel notification and managed VBA handoff; manual verification of
   Later, cancellation, restart, version number and database-open rescheduling.
2. Package upgrade: out-of-process handoff and signed cached installer can update
   XLAM, loader and updater together. Test actual old-to-new package transition,
   interruption, Excel reopening, reinstallation and complete recovery.
3. Automatic delivery: bounded HTTPS download, trusted signed feed, publisher
   verification, policy/opt-out/version pinning and background checks. Recheck
   policy and package identity immediately before applying. No caller-provided
   trust key or writable manifest may authorize arbitrary code.
4. Release candidate: clean standard-user machine, actual browser download,
   missing prerequisites, representative VBA data operations, diagnostic logs,
   licensing/release ownership and documented rollback/support.

0.6 is a test release, not production 1.0. Simulated VBA 6.02.006 and 6.02.007
must not be published as real NADABAS releases. No GitHub publishing or permanent
scheduled task installation is part of preparing these local test packages.

## Open release inputs

Production feed-signing key custody/rotation and exact approved release assets
must be established before unattended delivery is enabled. Existing GlobalSign
Authenticode signing is available; it does not replace the feed-signing design.
IT settings need an enforceable machine policy source, not a user-writable file.
The supported OS/Office matrix and installer licensing must be settled before
external distribution. Track evidence in PRODUCTION-READINESS.md.

## Candidate preparation, 2026-11-11

The maintainer selected the tested, signed VBA 6.02.009 payload, including
Exchange and issue 47, for the first delivery-platform release candidate.
Package 0.9.5 is the next internal candidate; 1.0 is not yet published.

`build_loader_test_installer.ps1 -CompactPayload` stages only baseline slots
slot-a/slot-b and slot-fix602009. Historical fixtures remain accepted for
upgrading existing installations but are not distributed in a fresh compact
package. Existing user files and previously installed slots are preserved.
The historical source arguments remain required for their verified hashes,
which the installer needs to recognize an older installation safely.

The compact option retains the current installation identity so the signed
0.9.4 updater can recognize the next package. Product renaming requires a
compatible bridge; changing the ProductName alone would break that check.
Compact packaging is not permission to publish a test-identity package as stable.

Next acceptance: build/sign the compact candidate, exercise fresh installation
and upgrade from the full test package, then prepare a reviewed feed. Final
publication still requires source approval, key custody and explicit approval
of the actual assets. The existing stable endpoint has no appcast asset.

The maintainer accepted responsibility for the feed-signing key on 2026-11-11.
The existing embedded public key remains unchanged. A recovery backup and drill
are still required; nominating a custodian does not establish those checks.
The current source HEAD was unsigned and no local allowed_signers path was
configured when checked. Production feed preparation must use the reviewed
source wrapper after these inputs are established.

Candidate evidence: 137 .NET tests and 129 Python tests passed; VBA security
scan reported no errors and 14 existing warnings. All pre-commit hooks passed
(the file-end hook required a successful retry after transient file access
failure), and .NET formatting checks passed. The compact signed installer is
6,674,424 bytes, compared with 14,886,264 bytes for 0.9.4. Fresh installation
passed 30 checks; full-to-compact 0.9.4 upgrade passed 40 checks, preserving
all nine historical payload files. The final guide-only rebuild has separate
installation evidence in artifacts/release-candidate-0.9.5/final-smoke-report.json.
Local feed/package/publisher verification passed for the final bytes. This is
not a successful GitHub download: the live channel preflight returned HTTP 404.
See artifacts/release-candidate-0.9.5/acceptance-summary.json for local records.

## Token-free preparation, 2026-11-12

The production candidate now uses the NADABAS display name while retaining the
AppId and ProductName required by installed 0.9.x updaters. `PrepareOnly` builds
without calling SignTool and blocks ordinary-profile installation. The isolated
variant passed 15 checks; signed-fixture package recovery passed 29 checks and
legacy migration passed 14. Source approval/SSH setup is the next maintainer
session, and token signing remains explicitly deferred. See V1-PREPARATION.md
and SSH-SETUP.md. External acceptance gates remain in PRODUCTION-READINESS.md.
