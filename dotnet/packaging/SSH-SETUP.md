# SSH source signing: maintainer walkthrough

This walkthrough concerns source signing. It changes
source approval only; it does not replace Windows/Office publisher signing or
the separate Ed25519 key that signs the update feed and package bytes.

Local acceptance, 2026-11-12: the maintainer created and registered a signing
key, enabled the Windows SSH agent and added the key under the ordinary user.
Repository-local signing and an external allowed_signers file are configured.
An empty commit in an isolated repository passed both git verify-commit and
tools/verify_release_source.py. The actual NADABAS preparation commit
7f100dfe1452cdbf123738ba2b18b042347d9cac also passed the approved-signer check.
This verifies local source signing only; GitHub verification and the release
environment still need acceptance. No existing history was rewritten and
nothing was pushed by these checks.

## 1. Inspect the existing setup

Check Git's version and repository-local signing settings. Git supports SSH
signature verification from 2.34; the current machine reported 2.53.0.windows.3.
No signing key, format or allowed_signers path was configured in the initial checks.
That does not mean no SSH key exists elsewhere.

List public `.pub` files in the user's SSH directory and inspect public-key
fingerprints only. Do not print private key files or collect passphrases.
Confirm whether an existing key belongs to this maintainer and is approved for
signing, or create a dedicated signing key outside the checkout with a local
passphrase prompt. Do not overwrite an existing key.

## 2. Choose the signing identity

Record the maintainer's Git identity, the chosen public-key path and SHA-256
fingerprint. Confirm the fingerprint with the maintainer before trusting it.
If using ssh-agent, ensure the selected private key is available to the same
Git/SSH toolchain that performs commits. The passphrase stays in the local
prompt. Agree protected storage and replacement/revocation responsibilities.

## 3. Register the public signing key on GitHub

Add the public key as a **Signing key** in the maintainer's GitHub account.
Authentication keys and signing keys have different roles even when the same
public key is used. This is an account change to perform in the setup session.
The private key never goes to GitHub or into the repository.

## 4. Configure this repository

After choosing and approving the real public-key path:

```powershell
git config --local gpg.format ssh
git config --local user.signingkey "C:/path/to/approved-signing-key.pub"
git config --local commit.gpgsign true
git config --local tag.gpgsign true
```

Use local configuration to avoid changing unrelated repositories. The GitHub
Desktop and terminal workflows must both be tested with the selected key.
GitHub documents SSH key selection in
[Telling Git about your signing key](https://docs.github.com/en/authentication/managing-commit-signature-verification/telling-git-about-your-signing-key).

## 5. Establish the approval list

Create an approved `allowed_signers` file outside the checkout. An entry maps
the approved principal to its public SSH key. Confirm the fingerprint through
the agreed review process; discovering a key alone is not approval. Configure
`gpg.ssh.allowedSignersFile` for convenient local verification. The release
verification tool also requires that file explicitly.

## 6. Verify a newly signed commit

Make the next reviewed commit with signing enabled, without rewriting existing
history just to add signatures. Verify it using:

```powershell
git verify-commit HEAD
python tools/verify_release_source.py --repository . --revision HEAD --allowed-signers "C:/approved/location/allowed_signers"
```

Verify the same commit on GitHub after an authorized push. GitHub's Verified
badge and NADABAS's approved-signer check answer different questions; both are
useful. The release tool rejects unsigned or unapproved source commits.

## 7. Connect release review

Configure the `nadabas-release-review` GitHub environment, approved public-key
variable and reviewers with the repository administrator. Confirm the merge
strategy preserves the required signed release commit. Only after source
review, artifact testing, publisher signing and key recovery checks should the
production feed be prepared. Publication remains a separate authorization.

See [Release signing and recovery](RELEASE-SIGNING.md) for the implemented
tools and [GitHub's existing-key guide](https://docs.github.com/en/authentication/connecting-to-github-with-ssh/checking-for-existing-ssh-keys)
for the initial public-key inventory.
