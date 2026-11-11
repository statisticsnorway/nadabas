"""Verify a release commit against an independently approved SSH signer file."""
import argparse
import json
import subprocess
from pathlib import Path


def verify(repository: Path, revision: str, allowed_signers: Path) -> dict:
    repository = repository.resolve(strict=True)
    allowed_signers = allowed_signers.resolve(strict=True)
    if allowed_signers.is_relative_to(repository):
        raise ValueError("Keep the approved signer file outside the source checkout.")
    if not allowed_signers.is_file() or allowed_signers.stat().st_size == 0:
        raise ValueError("An approved SSH allowed_signers file is required.")

    def git(*arguments):
        result = subprocess.run(
            ["git", "-C", str(repository), *arguments],
            check=True, capture_output=True, text=True,
        )
        return result.stdout.strip()

    commit = git("rev-parse", "--verify", "--end-of-options", revision + "^{commit}")
    signature = git("cat-file", "commit", commit).split("\n\n", 1)[0]
    if "gpgsig -----BEGIN SSH SIGNATURE-----" not in signature:
        raise ValueError("The release commit must carry an SSH signature.")
    git("-c", "gpg.format=ssh",
        "-c", "gpg.ssh.program=ssh-keygen",
        "-c", "gpg.ssh.allowedSignersFile=" + str(allowed_signers),
        "verify-commit", commit)
    return {"schemaVersion": 1, "commit": commit, "sshSignature": "verified"}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repository", type=Path, required=True)
    parser.add_argument("--revision", required=True)
    parser.add_argument("--allowed-signers", type=Path, required=True)
    args = parser.parse_args()
    try:
        print(json.dumps(verify(args.repository, args.revision, args.allowed_signers)))
    except (ValueError, OSError, subprocess.CalledProcessError) as error:
        parser.exit(1, f"Release source verification failed: {error}\n")
