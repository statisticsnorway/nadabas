import os
import struct

import pytest

from tools.prepare_runtime_bundle import inspect_runtime
from tools.prepare_runtime_bundle import prepare


def runtime(root, arch, version="10.0.11"):
    for relative, dll in (
        ("dotnet.exe", False),
        (f"host/fxr/{version}/hostfxr.dll", True),
        (f"shared/Microsoft.NETCore.App/{version}/coreclr.dll", True),
    ):
        path = root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        data = bytearray(512)
        data[:2] = b"MZ"
        struct.pack_into("<I", data, 60, 128)
        data[128:132] = b"PE\0\0"
        struct.pack_into("<H", data, 132, 0x14C if arch == "x86" else 0x8664)
        struct.pack_into("<H", data, 150, 2 | (0x2000 if dll else 0))
        struct.pack_into("<H", data, 152, 0x10B if arch == "x86" else 0x20B)
        path.write_bytes(data)
    desktop = (
        root / f"shared/Microsoft.WindowsDesktop.App/{version}/System.Windows.Forms.dll"
    )
    desktop.parent.mkdir(parents=True, exist_ok=True)
    desktop.write_bytes(b"managed fixture")
    return root


def test_prepares_both_architectures_with_one_manifest(tmp_path):
    x86 = runtime(tmp_path / "x86", "x86")
    x64 = runtime(tmp_path / "x64", "x64")
    output = tmp_path / "bundle"
    record = prepare(x86, x64, output)
    assert not record["canDistribute"]
    assert len((output / "runtime-manifest.tsv").read_text().splitlines()) == 8
    for arch in ("x86", "x64"):
        assert (
            inspect_runtime(output / "payloads" / arch / "runtime" / arch, arch)[1]
            == "10.0.11"
        )
    with pytest.raises(FileExistsError):
        prepare(x86, x64, output)


@pytest.mark.parametrize(
    "failure", ["wrong-arch", "missing-desktop", "extra", "mixed-version", "wrong-dll"]
)
def test_invalid_inputs_do_not_create_output(tmp_path, failure):
    x86 = runtime(tmp_path / "x86", "x64" if failure == "wrong-arch" else "x86")
    x64 = runtime(
        tmp_path / "x64", "x64", "10.0.12" if failure == "mixed-version" else "10.0.11"
    )
    if failure == "missing-desktop":
        (
            x86 / "shared/Microsoft.WindowsDesktop.App/10.0.11/System.Windows.Forms.dll"
        ).unlink()
    if failure == "extra":
        (x86 / "unexpected.dll").write_bytes(b"unapproved")
    if failure == "wrong-dll":
        (x86 / "dotnet.exe").write_bytes(
            (x86 / "host/fxr/10.0.11/hostfxr.dll").read_bytes()
        )
    with pytest.raises(ValueError):
        prepare(x86, x64, tmp_path / "bundle")
    assert not (tmp_path / "bundle").exists()


def test_output_cannot_be_inside_source(tmp_path):
    x86 = runtime(tmp_path / "x86", "x86")
    x64 = runtime(tmp_path / "x64", "x64")
    with pytest.raises(ValueError):
        prepare(x86, x64, x86 / "output")


@pytest.mark.skipif(os.name != "nt", reason="Windows packaging path limit")
def test_long_windows_output_fails_before_copying(tmp_path):
    x86 = runtime(tmp_path / "x86", "x86")
    x64 = runtime(tmp_path / "x64", "x64")
    output = tmp_path / ("long" * 45)
    with pytest.raises(ValueError, match="shorter output path"):
        prepare(x86, x64, output)
    assert not output.exists()


def test_symlink_input_is_rejected(tmp_path):
    target = runtime(tmp_path / "source", "x86")
    link = tmp_path / "link"
    try:
        link.symlink_to(target, target_is_directory=True)
    except OSError:
        pytest.skip("Creating symlinks is not permitted on this host")
    with pytest.raises(ValueError):
        inspect_runtime(link, "x86")
