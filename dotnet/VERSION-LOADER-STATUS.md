# Demo readiness, 2026-11-04

The signed compact 0.8.1 add-in was opened read-only in Excel. Its VBA language
setter, LoadLanguage and GetRibbon returned the expected French, Portuguese and
Bahasa Indonesia text. The previous per-user language preference was restored.
The four form button captions and their code mappings were checked separately
with macros disabled. This is runtime lookup evidence; the visible picker and
ribbon still need the short manual walkthrough in the demo guide.

All 942 resource rows containing English text have nonblank French, Portuguese
and Indonesian columns. Text quality and layout require visual review. The signed
package has English Exchange and .NET update controls. Source changes localize
those controls, remember the selection in the language picker and allow buttons
to size to their labels. The offline preview uses the same text catalog.
Demo guide, preview and evidence are in artifacts/demo-language.

Setup migration code compiled in a signing-disabled smoke build. No new signed
Setup migration run has been claimed. The 16 checks from the earlier rehearsal
describe the external harness around 0.8.1. A separate unsigned language XLAM
candidate was inspected with macros disabled. Signing and interactive
acceptance for those changes remain pending.

# Current work: 0.8 package-update engine candidate (2026-11-03)

User accepted retaining the conditional Exchange toggle and moving toward 0.8.
The toggle now depends on the active workbook's actual ExchLinks range and is
independent of the saved on/off preference. Signed beta slot-beta602006b retains
version 6.02.006; original beta slot remains available.

New signed Nadabas.PackageUpdate executable stages outside the installation,
verifies Authenticode/publisher/product/package version, and applies a cached
installer without running from the target directory. Package version is distinct
from the VBA version. Initial end-to-end 0.7.1 -> 0.8.0 isolated test passed 14
checks. Final verified-signer build passed all 17 expanded integration checks, including
installer rejection/repair-state logging. Signed candidate and evidence are in
artifacts/loader-installer/0.8.0-engine-candidate.
104 .NET tests and 20 Python/VBA/date/visibility checks passed.

This is NOT a completed automatic 0.8 user experience: whole-package readiness
and handoff are not yet connected to the Excel dialog; interrupted-installation
recovery needs further implementation/tests. See packaging/PACKAGE-UPDATER-0.8.md.
No GitHub download, scheduler or user-profile installation was enabled.

# Current package: Exchange beta 0.7.1 (2026-11-03)

User supplied a real unsigned beta, chose Exchange UI fix before packaging.
Separate signed candidate preserves version 6.02.006 with slot-beta602006;
originals and historical simulated fixtures are untouched. Default-off persisted
Show Exchange tools toggle gates all five groups, without changing connections.
99 .NET tests and 19 Python/VBA/date/visibility checks passed. Signed package
built; two-package isolated installer transition passed all 19 checks. Delivery: artifacts/loader-installer/0.7.1. Interactive
ribbon appearance and persistence remain pending. See packaging/EXCHANGE-BETA-REVIEW.md.

# 0.7.0 restart and build-date work (2026-11-03)

User reported 0.6.0 acceptance successful and supplied a screenshot of the
notification. Requested Restart and update and automatic release date stamping.
Source and signed fixture 6.02.008 prepared. Interactive restart acceptance
remains pending; do not infer it from the earlier 0.6.0 screenshot.

# Current milestone: signed 0.6.0 notification test (2026-11-03)

The isolated VBA 6.02.007 fixture is signed and admitted into the compiled
allowlist. Authenticode is Valid; macro-disabled read-only Excel inspection
confirmed VBASigned, the version constant, managed bridge and three guards.
The signed 0.6.0 installer has been built with this fixture and the Excel UI.
95 .NET tests pass. Interactive Excel notification acceptance remains pending.

The agreed 1.0 scope includes updating the updater and loader as well as VBA.
See packaging/V1-RELEASE-PLAN.md for the complete package handoff and release
gates. Automatic GitHub delivery, enforceable IT policy and full platform
upgrade/recovery are not implemented by this 0.6.0 package.

## Historical milestones

# Version loader prototype: local status, 2026-10-29

## Pending activation milestone, 2026-11-02

User reported 0.4.0 tests complete and selected GitHub Releases as the future
automatic-update source. In 0.5.0, queue/cancel-update CLI operations prepare a
pending activation, and the loader applies it at the next process startup.
Inspection does not apply it. Payloads are revalidated against the compiled
hash/version catalog, and failed requests leave the previous usable selection.
The runtime report includes updateStatus/updateError. Unit suite: 93 passed,
including queue timing, cancellation, explicit rollback, tampering, invalid/stale
pending records, downgrade refusal and crash-replay cleanup. Automatic GitHub
checks, download, user notification and production IT policy remain unimplemented.

The signed 0.5.0 package is delivered under artifacts/loader-installer/0.5.0.
The isolated package variant passed 17 lifecycle checks, including queueing
without immediate activation and cancellation through the installed CLI.
Format and pre-commit checks passed. Actual next-start activation in the user's
Excel session remains the next acceptance step; no GitHub release was published.

## Simulated release preparation, 2026-11-02

The maintainer reported 0.3.2 working and clarified that the next request is a
version-number-only upgrade simulation. `prepare_upgrade_test.py` copies only
the approved 6.02.005 artifact, changes CURRENT_VERSION to 6.02.006 and adds a
SIMULATED UPGRADE TEST About title, without importing unpublished repository VBA.
Saving that copy prompted Excel's token login. The first save finished without
a signature. Signing was retried successfully with SignTool and the installed
Office SIP: Authenticode reported Valid with the SSB certificate. A read-only
Excel inspection with macros disabled confirmed VBASigned, the 6.02.006 constant
and the About test label. The signed file SHA-256 is
852CC76872428186AC0F8B5C6AA906AE438DABEE6BEF8061E739858DCCD052B8.

The .NET release model now binds slots to both hashes and expected runtime
versions. 84 tests passed, including different old/new byte fixtures and rollback.
The compiled allowlist now admits that exact signed fixture in slot-test602006
alongside the original baseline slots. The packed loader was rebuilt and signed.
The new installer source targets 0.4.0 and requires an explicit signed upgrade
fixture plus UpgradeHash at compile time.
The finalized fixture's slot is to be slot-test602006; package activation selects
that release after installation. English guidance and prerequisite checks are
prepared. Manual acceptance of 0.4.0 remains required before external distribution.

The signed 0.4.0 package is delivered under artifacts/loader-installer/0.4.0.
Its isolated installer variant passed 15 lifecycle checks, including selection
of the simulated release, reinstall, preservation of reports, rejection of
tampered payload and uninstall cleanup. The .NET suite passed 84 tests, and
format/pre-commit checks passed. The original C:/NADABAS/NADABAS.XLAM hash remains
unchanged. No live Excel upgrade or database test is claimed for this package.

## Installer 0.3.2

Following the user's request, Setup now registers the signed session-only XLL
for next Excel startup and removes only exact legacy payload registrations in
its own test directory. Uninstall removes the same exact entries. All paths are
scoped to the installation; unrelated Office settings and add-ins are preserved.
The user no longer needs manual XLL activation or a close/reopen cycle to persist
that first manual activation. Runtime/security acceptance in their Excel profile
is still required; registry tests alone do not establish successful Excel loading.

The old fresh-directory requirement was replaced by validation of an existing
installation or recognized leftovers. The active pointer, slots and reports are
preserved. Existing slots are hash-checked, so reinstall does not silently replace
modified VBA files. Automatic updates remain a later milestone.

## Correction in 0.3.1 (manual acceptance pending)

The user confirmed reactivating the XLL, not a payload XLAM. Excel-DNA documents
that IExcelAddIn.AutoClose is not called on ordinary Excel shutdown. The 0.3.0
loader persisted the payload via AddIns.Installed and relied on AutoClose to undo
it; that assumption was incorrect. In 0.3.1, the payload is opened read-only via
Workbooks.Open without AddIns.Add/Installed. Explicit loader removal closes only
its owned workbook without saving. Normal process exit releases the file lease;
no persistent payload deregistration is required. Existing conflicting add-ins
still cause a safe refusal rather than being silently disabled.

The manual guide now requires closing all Excel processes after initial XLL
activation and verifying a normal restart before starting the two-process test.
This is a test prerequisite for Excel saving the activation, not proof that the
reported missing activation on /x is resolved. Reports include loaderVersion
0.3.1 and loadingMode. A loaded report is historical startup evidence; process
exit, not an AutoClose report, is the normal shutdown check.

Reference: https://excel-dna.net/docs/guides-advanced/detecting-excel-shutdown-and-autoclose/

## Manual follow-up

The user successfully activated the byte-identical copy placed under the shared
workspace `artifacts/manual-excel-test`. Their screenshot confirms the NADABAS
Ribbon and About version 6.02.005; they reported deactivation and Excel closure.
The earlier LocalAppData test directory was visible to the agent but not to the
user's Explorer. Filesystem isolation is therefore a material confounder in the
automated failures below, not evidence that the same loader fails in the user's
normal profile. The exact cause is still unconfirmed.

The new `packaging/loader-test.iss` package is for installation by the user in
their own Windows session, followed by the included manual guide. Manual XLAM
activation has passed; signed XLL startup, concurrent Excel sessions and rollback
within Excel have not yet passed acceptance.

## Earlier automated investigation

### User test of package 0.3.0

The user reported that starting a second process with `excel /x` required
reactivating the add-in, whereas ordinary startup appeared to work. Session
reports read from the installed test directory show `loaded` for PID 18708
(slot-a) and PID 28872 (slot-b). Later reports for PID 41668 (slot-b) and PID
17660 (slot-a) show `failed`: "Another NADABAS XLAM is active; loader did not
replace it." Therefore Ribbon visibility alone does not prove loader startup
or rollback. These snapshots do not establish simultaneous healthy sessions or
successful cleanup. Clarify whether the user reactivated the XLL or an XLAM.
The source currently persists payload activation through AddIns.Installed;
startup persistence and shutdown cleanup need investigation before acceptance.
Do not silently deactivate an unrelated NADABAS installation to bypass the guard.

The prototype selects an immutable XLAM slot through an atomic active pointer.
Each Excel session keeps its selected slot. Both test slots contain the approved
6.02.005 release, not distinct NADABAS releases. The Excel-DNA XLL has a valid
Authenticode signature in the isolated signing artifact.

Acceptance is **not passed**. After removing a stale startup entry from a prior
test, the XLL reported selecting slot-a, but Excel rejected `AddIns.Add`. Tests
with `/x /e` and `/x` both failed to load the payload. A separate COM test with
a blank workbook and without the version-store lease also rejected that file.
In the same COM session, the older successful startup-test copy was accepted.

The user explicitly authorized removing Internet-zone metadata from two test
copies. Only the two copies in
`NBLoader-efb6512779c54dac8a493e4496e07cdf` were successfully unblocked.
Their SHA-256 values remained the approved baseline. Source and normal installed
XLAM contents and zone metadata were verified unchanged. No Office trust settings
were changed. The script requires explicit `-UnblockTestCopies`; `-PreparedStore`
allows retrying these same files without unblocking additional copies.

The accepted older copy and failing slot have identical SHA-256 values, sizes,
ReadOnly/Archive attributes and no alternate data streams. Their inherited ACLs
differ: the older temporary copy has an additional local-account access entry.
Whether that difference explains the Excel failure has **not** been established.
Check effective access for the Excel process identity before changing the loader
or weakening any security control.

The failed-run report remains in `artifacts/version-loader/latest-report.json`.
After its cleanup timeout, the empty test Excel session was closed normally via
COM; an independent check confirmed no Excel process remained. The exact stale
test OPEN registration was then removed. The failed report was not rewritten
as a success. Concurrent sessions, version switching and rollback remain unverified.

0.6.0 final validation: 17 isolated installer lifecycle checks and 13 VBA checks passed. Signed delivery: artifacts/loader-installer/0.6.0. No live Excel acceptance claimed.

0.7.0 signed delivery: artifacts/loader-installer/0.7.0. Verified 99 .NET tests, 16 Python/VBA/date tests and 17 isolated installer lifecycle checks. Interactive Restart and update remains pending user acceptance.
