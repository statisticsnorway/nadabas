# Release signing and recovery

The release process verifies the source commit, signs the Windows deliverables,
and signs the feed and package bytes. The private release key belongs to the
release operator. End users receive public verification keys in the signed
updater. Keep the SSH development key and feed-signing key separate.

## SSH source approval

An operator maintains an OpenSSH allowed_signers file outside the checkout.
Each entry contains a principal and the approved public SSH key. Confirm the
fingerprint through an independently agreed channel before adding a key.
Obtain the signing public key from its owner or GitHub's SSH signing keys API.

Configure the developer repository using the approved signing key:

```powershell
git config --local gpg.format ssh
git config --local user.signingkey "C:/path/to/approved-signing-key.pub"
git config --local commit.gpgsign true
git config --local tag.gpgsign true
```

The private key must be available to OpenSSH, for example through ssh-agent.
Never put it in the checkout or send it in chat.

`tools/verify_release_source.py` verifies the exact commit against the external
allowed_signers file. Unsigned commits, signatures from other keys, and a trust
file inside the checkout are rejected. A GitHub web merge may produce a commit
with a different signature format or signer. Release the approved SSH-signed
commit and ensure the chosen merge strategy preserves that policy.

The manual **Verify release source** workflow has read-only repository access.
Configure the GitHub environment `nadabas-release-review` with reviewers and the
environment variable `NADABAS_RELEASE_ALLOWED_SIGNERS` containing approved public
key entries. Protect the workflow and environment configuration through the
repository's review policy. These remote settings require an administrator.
The workflow verifies the supplied full SHA and publishes nothing.

Use `tools/prepare_verified_feed.py` for release preparation. It requires a clean
SSH-signed HEAD, verifies its signer, invokes the existing publisher and feed
checks, and records the source SHA alongside the output. Build and review the
installer from that same commit. The source record supports traceability;
reproducible-build or CI artifact attestation is separate work.
The lower-level `prepare-feed` command remains available for local fixtures.

## Feed key custody and recovery

The existing DPAPI-protected working key remains the configured key. The
following commands are implemented by `Nadabas.FeedSigning`:

```text
backup-key <absolute-feed-key.dpapi> <new-absolute-backup-file>
restore-key <absolute-backup-file> <new-absolute-feed-key.dpapi> <expected-base64-public-key>
```

The tool prompts for a hidden recovery passphrase in the terminal. Backup asks
twice and requires at least 16 characters. Use a generated high-entropy
passphrase. The portable backup uses AES-256-GCM with a random salt and nonce,
PBKDF2-HMAC-SHA256 with 600,000 iterations, and a versioned authenticated context.
Restoration checks the independently recorded public key before writing a fresh
DPAPI working copy for the restoring Windows user. Existing files are refused.

SSB must designate a primary custodian and a recovery custodian. Store the
encrypted backup and recovery passphrase separately in approved storage. Record
the public key and its SHA-256 fingerprint in the release register. Exercise
recovery on the intended replacement workstation and verify a test signature
against the original public key before relying on the backup.

This work does not create or move the production backup, nominate custodians,
or claim a completed SSB recovery exercise. Automated recovery tests use
ephemeral keys.

## Planned rotation

`update-channel.json` retains the existing `PublicKey` property and additionally
supports `AdditionalPublicKeys` and `RevokedPublicKeys` arrays. Trust is embedded
in the signed updater. Canonical base64 encodings are required. Revocation takes
precedence over both primary and additional trusted keys.

1. Create the next key under the agreed custody process.
2. Ship a bridge updater containing the current key and the next key. Sign its
   feed and package with the current key so existing clients can install it.
3. Keep a feed signed with the current key available for clients that have not
   installed the bridge. A switch of the single shared feed to the new key
   immediately strands those older clients. Use a separate next-channel URL in
   the bridge, or complete the managed migration before switching the feed.
4. After bridge adoption, sign subsequent feed and package bytes with the new
   key. Remove or revoke the old key in the next signed updater.

Current clients with one embedded key cannot learn a new key from an arbitrary
downloaded file. Emergency revocation requires distribution of a trusted updater
through IT or the independently authenticated installer channel. An old offline
client cannot learn about revocation until its trusted configuration is updated.
No remote revocation service or runtime editable trust file is introduced.

## Acceptance before distribution

- Source commit verifies against the approved SSH signer policy.
- Recovery key produces the expected public key and a valid test signature.
- Feed rotation accepts the bridge keys and rejects revoked keys.
- Download verifies the signed feed, approved HTTPS redirects, and package bytes.
- Windows publisher verification, Excel startup, restart, unsaved work,
  other Excel sessions, upgrade and explicit recovery are exercised.

The TLS transport test runs the actual HTTP client against an ephemeral local
TLS server while retaining the approved logical GitHub URLs. Certificate pinning
and address routing exist only in the test. The test downloads fixture bytes.
An actual GitHub release download and Office acceptance require signed release
artifacts and remain separate acceptance steps.

For desktop acceptance launched from Codex, use an ordinary Windows session
for staging and installation and verify the physical cache location. The
packaged Codex process can see a different LocalCache and registry view.

References:

- https://github.com/statisticsnorway/statbus/blob/master/doc/SECURITY.md
- https://github.com/statisticsnorway/statbus/blob/master/cli/cmd/upgrade.go
- https://github.com/NetSparkleUpdater/NetSparkle
