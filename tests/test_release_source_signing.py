import subprocess
import shutil
from pathlib import Path
import sys

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))
from verify_release_source import verify


@pytest.mark.skipif(not shutil.which("ssh-keygen") or not shutil.which("git"), reason="Git and OpenSSH required")
def test_release_requires_approved_ssh_signature(tmp_path):
    repository = tmp_path / "repo"
    repository.mkdir()
    key = tmp_path / "ephemeral-key"
    subprocess.run(["ssh-keygen", "-q", "-t", "ed25519", "-N", "", "-f", str(key)], check=True)
    allowed = tmp_path / "allowed_signers"
    allowed.write_text("test-signer " + key.with_suffix(".pub").read_text(), encoding="utf-8")

    def git(*args):
        return subprocess.run(["git", "-C", str(repository), *args], check=True, capture_output=True, text=True)

    git("init")
    git("config", "user.name", "Ephemeral Test")
    git("config", "user.email", "test@example.invalid")
    git("config", "gpg.format", "ssh")
    git("config", "user.signingkey", str(key))
    git("commit", "--allow-empty", "--no-gpg-sign", "-m", "Unsigned fixture")
    with pytest.raises(ValueError):
        verify(repository, "HEAD", allowed)
    git("commit", "--allow-empty", "-S", "-m", "Signed fixture")
    assert verify(repository, "HEAD", allowed)["sshSignature"] == "verified"
    with pytest.raises((ValueError, subprocess.CalledProcessError)):
        verify(repository, "--help", allowed)
    copied_policy = repository / "allowed_signers"
    copied_policy.write_text(allowed.read_text(), encoding="utf-8")
    with pytest.raises(ValueError):
        verify(repository, "HEAD", copied_policy)
    other_key = tmp_path / "other-key"
    subprocess.run(["ssh-keygen", "-q", "-t", "ed25519", "-N", "", "-f", str(other_key)], check=True)
    allowed.write_text("other " + other_key.with_suffix(".pub").read_text(), encoding="utf-8")
    with pytest.raises(subprocess.CalledProcessError):
        verify(repository, "HEAD", allowed)
