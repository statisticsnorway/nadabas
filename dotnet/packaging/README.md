# Installation prototype: Inno Setup + NetSparkle

This is a disposable Windows x64 installation experiment, not a NADABAS release.
Package versions **0.1.0** and **0.2.0** both contain the same approved, unchanged
NADABAS **6.02.005** XLAM. The independent package version lets us test updates
without modifying the signed VBA artifact or inventing a newer NADABAS release.

## What is being evaluated

- Inno Setup **7.1.0** builds a normal per-user EXE installer and uninstaller.
- NetSparkle **3.1.0** discovers an update from a signed appcast. Its strict
  Ed25519 verifier also checks downloaded package bytes. The probe uses HttpClient
  for the package transfer and never executes the result; the acceptance harness
  runs the verified installer as a separate, explicitly bounded test step.
- Both .NET executables ship self-contained for win-x64. This avoids an end-user
  SDK requirement, but still needs validation on a clean machine without .NET.
- The installer has a separate AppId, name and `%LOCALAPPDATA%` directory. It has
  no Excel registry changes, scheduled tasks, service, Start Menu entries or
  process-termination actions. Unknown user files are not removed on uninstall.
- Excel acceptance uses the installed diagnostic executable to activate a fresh
  disposable copy of the packaged XLAM, then remove that test registration. It
  does **not** establish persistent loading from the installed version directory.

NetSparkle is an updater library, not an installer builder. Inno Setup was chosen
for the first experiment because it directly produces an EXE and provides silent
install/uninstall. WiX remains an option for MSI/managed enterprise deployment;
it was not implemented or benchmarked in this experiment.

## Reproduce on Windows

Use desktop Excel, .NET 10 SDK, the project's Python environment (only for a
loopback HTTP server), and the approved release at `C:\NADABAS\NADABAS.XLAM`.
Save and close Excel first. Run in PowerShell from the repository root:

```powershell
./tools/prepare_installer_prototype.ps1
./tools/build_installer_prototype.ps1 -Compiler "$PWD/artifacts/installer-prototype/tooling/inno/ISCC.exe"
./tools/test_installer_prototype.ps1
```

Preparation downloads a pinned official compiler, checks its published SHA-256
and Authenticode validity, and extracts it in portable mode below `artifacts/`.
NetSparkle's appcast generator **2.9.0** is installed to that same local tool area.
No compiler file associations or global dotnet tools are installed.

The harness installs to `%LOCALAPPDATA%\NADABAS-Installation-Prototype` and creates
only the corresponding test program's current-user uninstall registration. It
refuses an existing destination or registration, and does not touch a running
Excel session. Results and logs are written below
`artifacts/installer-prototype/run-<id>/`; `latest-report.json` summarizes the last
run. The loopback server is stopped and the prototype uninstalled on normal
completion. Cleanup is also attempted on failure. Inspect any reported cleanup
error before rerunning. The `user-data/preserve-me.txt` sentinel intentionally
remains to demonstrate data preservation; move that test directory aside before
rerunning. Do not remove unrelated files to make a rerun pass.

The test checks Windows registration, installed executable versions, packaged
XLAM activation before/after upgrade, genuine signed-feed acceptance, rejection
of a modified feed and modified package, no download on rejection, an up-to-date
result after upgrade, uninstall and preservation of a user-data sentinel. It
also hashes the original release and normal installed add-in before/after.

## Local result, 2026-10-28

The local run `run-68a2ef36c7474bb89abb4b537864e802` passed **18 acceptance
checks**. Both installed package versions ran, and Excel 16.0 activated the
disposable copy and reported NADABAS 6.02.005 before and after upgrade. Both
modified-feed and modified-package tests were rejected. The prototype was
uninstalled and its Windows registration removed. The test user-data sentinel
was retained; the normal installed XLAM and release source hashes were unchanged.
The .NET suite also passed all 74 existing tests.

The first attempt failed at Excel `AddIns.Add` with the long test artifact path.
The harness now places Excel copies in a shorter, unique `%TEMP%\NB-<id>` path;
the rerun passed. No Office trust or Windows long-path settings were changed.
This identifies short installation/test paths as an acceptance concern, not a
verified universal maximum-path guarantee. Full evidence and installer logs are
in ignored artifacts; no test signing key belongs in source control.

## Boundaries and release requirements

Follow-up on 2026-10-29: the Excel smoke-test transaction copies file contents
through a stream and does not preserve `Zone.Identifier`. The approved source
has Internet-zone metadata; the successful smoke-test copy did not. Those results
therefore verify activation of the copied bytes, not activation of a downloaded
file with its original security metadata. No Office security settings were changed.
Production installation needs an explicit provenance and trust design; the smoke
test must not be treated as evidence that downloaded add-ins will activate.

The probe accepts **only literal 127.0.0.1 HTTP URLs**. The public key is a local
test argument; disposable private keys stay in ignored artifacts and must never
become production keys. This is not a trusted production distribution channel.
The EXE installers are not Authenticode-signed. An Ed25519 test signature does
not establish an Office or Windows trusted publisher.

Downloads are held in memory and verified before writing; production needs
bounded streaming, pinned publisher trust, secure transport, policy-controlled
sources and version selection. Automatic scheduling, IT policy enforcement,
installation recovery, downgrade control and version retention are not provided
by this probe. The installer does not solve the Excel startup/activation race.
Database regression, Ribbon observation, persistent activation and multi-session
loading remain separate tests. Current-profile success is not clean-VM evidence.

Before production, evaluate who owns updates (user updater or IT), Authenticode
and Office signing, and the chosen installer's current license/support terms.

## Primary references

- [NetSparkle 3.1.0](https://github.com/NetSparkleUpdater/NetSparkle/releases/tag/3.1.0)
- [NetSparkle documentation and installer boundary](https://github.com/NetSparkleUpdater/NetSparkle/tree/3.1.0)
- [Inno Setup 7.1.0](https://github.com/jrsoftware/issrc/releases/tag/is-7_1_0)
- [Inno Setup portable extraction source](https://github.com/jrsoftware/issrc/blob/is-7_1_0/setup.iss)
- [Inno Setup commercial-use information](https://jrsoftware.org/isorder.php)
- [WiX maintenance terms](https://docs.firegiant.com/wix/osmf/)
