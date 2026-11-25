# 0.8 package-update engine candidate

Implemented: a signed self-contained .NET executable prepares a separate cache
outside the installed directory. It copies itself and a verified signed installer
there, checks the registered per-user installation, and requires a strictly newer
PACKAGE version. VBA version is independent; the retained beta still reports 006.
The cached helper can replace installed updater/loader files without locking them.

Preparation and execution verify Windows Authenticode trust and the approved
publisher. The publisher certificate is taken from the verified Windows trust
state. Product identity distinguishes normal and isolated smoke installers.
SHA-256 pins the staged bytes, and a read lease prevents replacing the installer
between final verification and execution. Windows trust verification may need
network access for revocation checks; failure stops the operation.

References:
- https://learn.microsoft.com/windows/win32/api/wintrust/nf-wintrust-winverifytrust
- https://learn.microsoft.com/windows/win32/api/wintrust/nf-wintrust-wthelpergetprovsignerfromchain

The helper refuses installation while Excel is running and checks again before
starting Setup. Setup retains its own Excel guards. No Excel process is killed.
An exclusive job lock prevents concurrent application of the same job. Once
execution has started, automatic replay is refused. Exit codes and post-install
registry version are checked; failure records repair-required plus an installer
log. This is not transactional recovery of every possible partially replaced file.

Developer-only commands (use absolute paths):

    Nadabas.PackageUpdate.exe stage <installed-root> <signed-installer> <new-cache>
    <new-cache>/Nadabas.PackageUpdate.exe apply <new-cache>

For tests use only /DSmokeTest installers and add --smoke to stage. The test script
refuses normal installers and pre-existing smoke registry state, then installs the
old package, stages the next package, checks tampering/replay/failure cases, upgrades
from outside the installation and uninstalls the isolated instance.

## Remaining before calling the complete 0.8 experience finished

- Connect verified whole-package readiness and handoff to the Excel notification;
  the current Excel dialog still operates on XLAM slots only.
- Wait/cancel handoff after normal Excel shutdown and test user-visible failures.
- Verify full platform recovery for interrupted/partially applied installations;
  this candidate records repair-required and preserves diagnostics instead of
  claiming automatic rollback.
- Confirm installation and conditional Exchange toggle manually in Excel.

GitHub feed/download integration, policy enforcement and production key rotation
remain later release gates. No production feed, scheduled task or live update
was enabled by these local tests. The conditional Exchange toggle is retained:
it is visible only for an active workbook with a valid ExchLinks range, independent
of whether the user has chosen to show the Exchange groups.

Final validation (2026-11-03): 104 .NET tests, 20 Python/VBA/date/visibility
checks, and 17 signed-package integration checks passed. The final helper rejects
a different publisher; the whole-package test rejects tampered cache bytes, logs
installer rejection, preserves the old loader on rejection, upgrades 0.7.1 to
0.8.0 successfully and rejects replay. This is isolated registry evidence, not
live Excel UI acceptance or full interrupted-installation rollback evidence.
