"""Prepare an update feed only from a clean checkout with an approved SSH commit."""
import argparse
import json
import subprocess
from pathlib import Path

from verify_release_source import verify


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--allowed-signers", type=Path, required=True)
    parser.add_argument("--feed-key", type=Path, required=True)
    parser.add_argument("--installer", type=Path, required=True)
    parser.add_argument("--tag", required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    repository = Path(__file__).resolve().parents[1]
    status = subprocess.check_output(
        ["git", "-C", str(repository), "status", "--porcelain"], text=True,
    )
    if status.strip():
        raise ValueError("Commit and review the release changes before preparing a feed.")
    evidence = verify(repository, "HEAD", args.allowed_signers)
    output = args.output.resolve()
    if output.exists():
        raise ValueError("Choose a new output directory.")
    subprocess.run([
        "dotnet", "run", "--project", str(repository / "dotnet/tools/Nadabas.FeedSigning"),
        "-c", "Release", "--", "prepare-feed", str(args.feed_key.resolve(strict=True)),
        str(args.installer.resolve(strict=True)), args.tag, str(output),
    ], check=True)
    (output / "source-verification.json").write_text(json.dumps(evidence, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
