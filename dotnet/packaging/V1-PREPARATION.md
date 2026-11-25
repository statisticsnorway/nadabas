# Preparing delivery 1.0 without the token

The maintainer requested token signing only after the other preparation and
tests, on 2026-11-11. This procedure creates engineering artifacts, not an
approved release. VBA 6.02.009 is the selected payload; its existing signature
and immutable hash are retained.

## Standard-user installation gate

The maintainer requires installation without administrator credentials, including
when the shared modern .NET runtime is absent. See [the local-runtime
investigation](LOCAL-RUNTIME.md). A manual prerequisite download does not satisfy
this gate. Existing prepared packages are not yet accepted for this scenario.

## Build and identity

Build the Excel bootstrap and publish the framework-dependent updater with
`-p:Version=1.0.0`. Pass those files to `build_loader_test_installer.ps1` with
`-ProductionCandidate -PrepareOnly -PackageVersion 1.0.0` and the existing
verified VBA inputs. `SignTool` is not required or invoked in this mode.
The historical `SignedLoader`/`SignedPackageUpdater` parameter names are retained
for compatibility; preparation mode accepts unsigned component inputs.

The output is `UNSIGNED-Nadabas-1.0.0-Setup.exe`. It refuses installation into
the ordinary Office profile. The isolated `/DSmokeTest` variant may install
for engineering tests under synthetic registration. Neither path disables
Office trust or the updater's publisher verification.

The production profile displays **NADABAS** and includes the user guide instead
of the prototype test guide. It ships only baseline slots a/b and the approved
6.02.009 slot. Old slots on existing installations are preserved.

AppId, internal installation directory, loader filename and version-resource
ProductName remain compatible with 0.9.x. The old updater pins ProductName to
`NADABAS Loader Test`, so that metadata must not be renamed independently.
The visible AppName is separate; this avoids a duplicate installation or a
broken upgrade path. See Inno's
[VersionInfoProductName documentation](https://jrsoftware.org/ishelp/topic_setup_versioninfoproductname.htm).

## Verification before signing

Run `tools/test_unsigned_installer.ps1 -BuildDirectory <prepared-build>
-Compiler <ISCC.exe>`. It compiles a separate unsigned smoke package, checks
that ordinary installation is blocked, then checks isolated installation,
all four language selections, version selection, compact contents, the guide,
registration and uninstall. It does not establish Office acceptance.

Use already signed 0.9.4/0.9.5 fixtures to exercise the unchanged publisher,
staging and recovery paths without the token. `test_package_update.ps1` takes
explicit `ExpectedSlot` and `ExpectedVbaVersion` inputs. The test now exercises
rejection of altered packages, installation failure, refusal to replay a failed
job, rejection of the wrong recovery package, actual restoration of the
previous package, successful upgrade and preservation of all previous XLAMs.
`test_legacy_migration.ps1` takes the same expected-payload inputs.

## Final signing handoff

1. Review the diff and establish the approved SSH-signed source commit using
   [the SSH walkthrough](SSH-SETUP.md). Rebuild from that reviewed source.
2. Retain the prepared build record and hashes. Sign the final loader and
   updater copies with the SSB token; verify their publisher signatures.
3. Re-run the builder with those signed files, `-ProductionCandidate`, the same
   package version and `SignTool`, **without** `PrepareOnly`. Use a new output
   directory. The builder signs the uninstaller and installer and verifies them.
   Do not merely sign the unsigned engineering EXE: its install-blocking guard
   is compiled in and its uninstaller is unsigned.
4. Record final hashes, repeat signed installation/upgrade and ordinary Excel
   acceptance, then prepare the feed with `prepare_verified_feed.py`.
5. Obtain approval for the exact public assets/channel before publishing.
   Exercise actual GitHub discovery through Excel restart afterwards.

New preparations also write local `build-inputs.json` for repeating the build.
It contains paths and a public certificate thumbprint, not keys. Review those
inputs and set PrepareOnly to false only at the authorized signing step.
`sourceCommit` and `sourceHasLocalChanges` in the build record provide context;
they are not SSH verification or reproducible-build attestation.

## Remaining external acceptance

Token signing is deliberately deferred. Clean standard-user Windows/Office
acceptance without an SDK, organizational publisher policy, representative
database regression, feed-key recovery custody, installer licensing and actual
GitHub distribution still require evidence. An unsigned smoke test cannot close
those gates. Keep the detailed matrix in PRODUCTION-READINESS.md.
