# Automatic downloads and Excel restart candidate (0.9.0)

The agreed repository is `statisticsnorway/nadabas`. The signed updater embeds
the HTTPS appcast URL and Ed25519 trust keys in `update-channel.json` as a resource.
It does not accept feed addresses or verification keys from command-line arguments,
the registry or an adjacent writable configuration file.

## User flow

Excel starts a separate signed update checker in the background. It checks once
on startup, then at most once every six hours across Excel processes. The NADABAS
Check Updates command can check immediately. There is no scheduled task or
service, so checks require Excel to be running. The checker never installs.

The appcast and installer each need an Ed25519 signature. The installer also needs
the pinned SSB Authenticode publisher and expected product/version metadata.
Downloads have a three-minute overall deadline, a 1 MiB feed limit and a 128 MiB
package limit. Redirects are limited to the approved repository and GitHub release
asset host. No HTTP downgrade, arbitrary host, userinfo or alternative port is
accepted. Verification uses NetSparkle's strict Ed25519 checker and appcast format.

Only a verified, staged update triggers the Excel ready message. Restart and
update checks for unsaved work and other Excel processes, waits for the original
Excel process to exit, applies the signed installer from an external cache and
reopens Excel after successful installation. It never force-closes Excel or
automatically reopens user workbooks. Cancellation of Excel shutdown times out
after 60 seconds and preserves the pending update for another attempt.

Later dismisses the notice for the session. Cancel records the skipped package
version, so automatic checks do not immediately offer it again. A manual check
may offer that version again. Failed installation retains logs and requires
review/repair; it is not silently replayed or described as an automatic rollback.

## IT policy and user preference

Machine policy is read from the **64-bit** registry view at
`HKLM\Software\Policies\NADABAS\Updates`.

| Value | Type | Meaning |
|---|---|---|
| Enabled | DWORD | 0 disables package updates, including manually requested updates. Default 1. |
| Automatic | DWORD | 0 disables background discovery/downloads. Manual requests remain available. Default 1. |
| PinnedVersion | REG_SZ | Only this exact normalized package version may be offered or applied. Downgrades are still refused. |

Users can opt out of background checks through `Automatic=0` (DWORD) at
`HKCU\Software\NADABAS\Updates`. User settings cannot override a machine deny.
Policy is re-read before download, staging, restart verification and installation.
Explicit administrator repair/recovery is separate from automatic upgrade policy.

## Release key and publishing

See [Release signing and recovery](RELEASE-SIGNING.md) for SSH source verification,
portable encrypted recovery, overlapping keys, revocation, and the release
preparation wrapper. The current embedded key is preserved.

`dotnet/tools/Nadabas.FeedSigning` is a release-workstation tool, never shipped in
the installer. `new-key` creates a random Ed25519 seed encrypted with Windows DPAPI
CurrentUser under `%LOCALAPPDATA%\NADABAS-ReleaseSigning\<id>`. The private seed is
never written as plaintext or printed. The public key is pinned at build time.
The local key location is recorded in `artifacts/automatic-update-tests/feed-key-location.json`.

DPAPI is tied to this Windows user context. The release tool now provides
`backup-key` and `restore-key` with a hidden recovery passphrase. Before production,
assign IT custody and exercise recovery on the replacement workstation. Key
rotation requires a signed bridge release that trusts the next public key before
the old key is retired, with a feed still available to clients awaiting the bridge.

After building and signing the final installer, run the release tool with
`prepare-feed <absolute-feed-key.dpapi> <absolute-signed-installer> <release-tag>
<new-absolute-output-directory>`. It verifies SSB Authenticode and creates the
package, `appcast.xml`, `appcast.xml.signature`, and a review record. It uploads
nothing. Public GitHub publication is a separate reviewed action.

For production, the release assets must match the immutable tag in the appcast.
Upload all three assets together to a GitHub Release only after acceptance. The
current discovery address is `https://github.com/statisticsnorway/nadabas/releases/latest/download/appcast.xml`.
Do not mark internal demo or simulated VBA fixtures as a public latest release.
The missing/unpublished feed is handled as an unavailable check, with no changes
to the installation. No feed has been published by this work.

## Acceptance evidence and remaining work

All 125 .NET tests passed. Signed-feed tests cover altered data, wrong keys,
redirects, size limits, cancellation and version pins using a controlled HTTP
transport. The signed 0.8.2 to 0.9.0 isolated installer upgrade passed 31
lifecycle checks. The signed external helper installed 0.9.1 from 0.9.0 after
an actual Excel exit, preserved the language and reopened Excel. The user saw
the Excel window and closed it normally. The build, lifecycle and restart
records accompany the local 0.9.0 demo installer.

The **Restart and update** button was exercised in normal Excel with the
corrected 0.9.2 loader. After staging from the ordinary desktop, the job
reported package 0.9.3 installed, the pointer cleared, and the reopened Excel
process reported VBA 6.02.006 loaded without an error. Both test packages use
loader 0.9.2. The desktop job recorded its starting Windows registration as
0.8.2, so this run verifies restart and installation rather than a clean
registered 0.9.2-to-0.9.3 sequence. The package was staged locally.
Also check an unsaved workbook and a second Excel process.
Discovery through GitHub requires a published approved release and a
signed appcast at the configured URL. The prepared demo feed is local.

For a normal-profile acceptance run, launch the baseline Setup and the manual
`stage-for-excel` command from the user's ordinary Windows desktop session.
Use that session for `inspect-for-excel` and post-install checks as well.
In the 0.9.2 test, commands launched from the packaged Codex app saw a private
LocalCache directory and a different uninstall registration. Excel could read
the pending pointer but could not find the selected job's `package-update.json`.
Check the job's physical file location and its recorded previous version before
retrying. Leave package and publisher verification enabled.

Reference: [NetSparkle appcast and strict signature requirements](https://github.com/NetSparkleUpdater/NetSparkle#app-cast).

## Update failure details

The Excel notification offers **Details** after a failed manual verification or
update check. Background failures remain quiet. The report can be selected and
copied for IT support. Restart failures display the same diagnostic view.
All four supported languages provide the surrounding instructions.

Reports contain a random reference, UTC time, operation stage, updater assembly
version and up to eight nested exception types/HRESULTs. They intentionally omit
exception messages, stack traces, URLs, file paths and exception Data. Error
codes identify a failure category; they do not by themselves establish its cause.
The checker returns this structured failure to Excel instead of losing it on
standard error. Existing publisher and signature checks remain in force.

The installation's `diagnostics` directory retains `last-download.json`,
`last-verification.json` and `last-restart.json`, replacing the previous report
for each stage. Concurrent sessions may replace a report; the displayed report
retains its own reference. If writing fails, Details still works from memory.
Reports are not uploaded automatically. This change does not sanitize older
installer/session logs; review those separately before sharing them.

Visual validation pending: trigger a verification failure in an isolated signed
installation, use Details with keyboard navigation in each language and at larger
display scales, then retry with a valid package. Confirm the normal restart offer
and installed-version checks still work. Automated report tests are not Excel UI
acceptance.
