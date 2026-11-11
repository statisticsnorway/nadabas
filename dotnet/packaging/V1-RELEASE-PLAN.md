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
