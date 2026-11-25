"""Prepare architecture-checked runtime inputs for an unsigned engineering build."""

import argparse
import hashlib
import json
import os
import re
import shutil
import stat
import struct
from pathlib import Path


def plain_path(path):
    path = Path(os.path.abspath(path))
    for ancestor in (path, *path.parents):
        if ancestor.is_symlink() or (
            ancestor.exists()
            and getattr(ancestor.lstat(), "st_file_attributes", 0)
            & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0)
        ):
            raise ValueError("Filesystem links are not supported")
    return path


def digest(path):
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest().upper()


def check_image(path, architecture, dll):
    with path.open("rb") as stream:
        header = stream.read(64)
        if len(header) != 64 or header[:2] != b"MZ":
            raise ValueError(f"Invalid executable: {path}")
        offset = struct.unpack_from("<I", header, 60)[0]
        if offset < 64 or offset > path.stat().st_size - 26:
            raise ValueError(f"Invalid PE offset: {path}")
        stream.seek(offset)
        pe = stream.read(26)
    machine, magic = (0x14C, 0x10B) if architecture == "x86" else (0x8664, 0x20B)
    if (
        pe[:4] != b"PE\0\0"
        or struct.unpack_from("<H", pe, 4)[0] != machine
        or struct.unpack_from("<H", pe, 24)[0] != magic
        or bool(struct.unpack_from("<H", pe, 22)[0] & 0x2000) != dll
        or not struct.unpack_from("<H", pe, 22)[0] & 2
    ):
        raise ValueError(f"Expected {architecture} {'DLL' if dll else 'EXE'}: {path}")


def inspect_runtime(root, architecture):
    root = plain_path(root)
    files = {}
    seen = set()
    for directory, folders, names in os.walk(root, followlinks=False):
        for name in folders + names:
            plain_path(Path(directory) / name)
        for name in names:
            path = Path(directory) / name
            relative = path.relative_to(root).as_posix()
            if any(ord(c) < 32 or c in "\\:" for c in relative) or any(
                part.endswith((".", " ")) for part in relative.split("/")
            ):
                raise ValueError("Invalid portable runtime path")
            if relative.lower() in seen:
                raise ValueError("Case-insensitive duplicate runtime path")
            seen.add(relative.lower())
            files[relative] = digest(path)
    hosts = [p for p in files if re.fullmatch(r"host/fxr/10\.0\.\d+/hostfxr\.dll", p)]
    if len(hosts) != 1:
        raise ValueError("Exactly one stable .NET 10 host is required")
    version = hosts[0].split("/")[2]
    core = f"shared/Microsoft.NETCore.App/{version}/"
    desktop = f"shared/Microsoft.WindowsDesktop.App/{version}/"
    allowed = (f"host/fxr/{version}/", core, desktop)
    for path in files:
        if path != "dotnet.exe" and not path.startswith(allowed):
            raise ValueError(f"Unexpected runtime file or version: {path}")
    for path in (
        "dotnet.exe",
        core + "coreclr.dll",
        desktop + "System.Windows.Forms.dll",
    ):
        if path not in files:
            raise ValueError(f"Incomplete Desktop Runtime: {path}")
    for path, dll in (
        ("dotnet.exe", False),
        (hosts[0], True),
        (core + "coreclr.dll", True),
    ):
        check_image(root / path, architecture, dll)
    return root, version, files


def prepare(x86, x64, output):
    sources = {
        arch: inspect_runtime(path, arch) for arch, path in (("x86", x86), ("x64", x64))
    }
    if sources["x86"][1] != sources["x64"][1]:
        raise ValueError("Use the same reviewed runtime version for both architectures")
    output = plain_path(output)
    for source, _, _ in sources.values():
        if output == source or source in output.parents or output in source.parents:
            raise ValueError("Output and source directories must not overlap")
    if os.name == "nt" and any(
        len(str(output / "payloads" / arch / "runtime" / arch / relative)) >= 248
        for arch, (_, _, files) in sources.items()
        for relative in files
    ):
        raise ValueError(
            "Choose a shorter output path, for example a new directory under TEMP"
        )
    output.mkdir(parents=True, exist_ok=False)
    lines = []
    inventory = []
    for arch, (source, version, files) in sources.items():
        target = output / "payloads" / arch / "runtime" / arch
        size = 0
        for relative, expected in sorted(files.items()):
            original = plain_path(source / relative)
            destination = plain_path(target / relative)
            destination.parent.mkdir(parents=True, exist_ok=True)
            with original.open("rb") as src, destination.open("xb") as dst:
                shutil.copyfileobj(src, dst)
            if digest(destination) != expected:
                raise ValueError(
                    "Runtime changed during preparation; output is incomplete"
                )
            size += destination.stat().st_size
            lines.append(f"{arch}\t{relative}\t{expected}")
        # Recheck native headers and complete destination inventory before marking ready.
        if inspect_runtime(target, arch)[2] != files:
            raise ValueError("Prepared runtime differs from reviewed input")
        inventory.append(
            {
                "architecture": arch,
                "runtimeVersion": version,
                "files": len(files),
                "bytes": size,
            }
        )
    manifest = output / "runtime-manifest.tsv"
    manifest.write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")
    record = {
        "schemaVersion": 1,
        "status": "engineering-runtime-inputs-prepared",
        "canDistribute": False,
        "runtimeManifestSha256": digest(manifest),
        "architectures": inventory,
        "remaining": [
            "Runtime provenance and redistribution files review",
            "Component build and signing",
            "Installer architecture selection and desktop acceptance",
        ],
    }
    (output / "bundle-record.json").write_text(
        json.dumps(record, indent=2) + "\n", encoding="utf-8"
    )
    return record


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--runtime-x86", type=Path, required=True)
    parser.add_argument("--runtime-x64", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    print(
        json.dumps(prepare(args.runtime_x86, args.runtime_x64, args.output), indent=2)
    )


if __name__ == "__main__":
    main()
