"""Create an isolated, unsigned version-number-only VBA upgrade fixture."""

import argparse
import hashlib
import json
import re
import shutil
from datetime import date
from pathlib import Path

BASELINE = "2F7432B9B04C414337EB1AFEB07E2B93F05E4A5A6D151073FA967A2F7EA89A85"


def managed_bridge(code):
    """Apply only the reviewed update handoff, preserving other released VBA."""
    reviewed = (
        Path(__file__).resolve().parents[1] / "vba/modules/InterfaceVersionUpdate.bas"
    ).read_text(encoding="utf-8-sig")
    begin = "' BEGIN MANAGED UPDATE BRIDGE"
    end = "' END MANAGED UPDATE BRIDGE"
    bridge = reviewed[reviewed.index(begin) : reviewed.index(end) + len(end)]
    declaration = "Private checkIsScheduled As Boolean"
    if code.count(declaration) != 1 or "EnableManagedUpdates" in code:
        raise RuntimeError("Unexpected baseline update module")
    code = code.replace(
        declaration, declaration + "\nPrivate managedUpdates As Boolean"
    )
    for name, gate in (
        (
            "ScheduleLatestVersionCheck",
            "    If managedUpdates Then\n        CancelScheduledVersionCheck\n        Exit Sub\n    End If\n",
        ),
        ("CheckLatestVersion", "    If managedUpdates Then Exit Sub\n"),
        (
            "CheckLatestVersionManually",
            "    If managedUpdates Then\n        ShowManagedUpdateStatus\n        Exit Sub\n    End If\n",
        ),
    ):
        pattern = rf"(Public Sub {name}\([^\n]*\n.*?On Error GoTo [^\n]*\n)"
        code, count = re.subn(
            pattern, lambda m: m[1] + gate, code, count=1, flags=re.DOTALL
        )
        if count != 1:
            raise RuntimeError("Missing baseline procedure: " + name)
    return code + "\n" + bridge + "\n"


def stamp_build_date(code, about_code, build_date):
    # Validate ISO input; metadata must be immutable before VBA signing.
    iso = date.fromisoformat(build_date).isoformat()
    if "RELEASE_DATE_ISO" in code or "AboutBuildDate" in code:
        raise RuntimeError("Baseline already contains build metadata")
    version_pattern = r'(Private Const CURRENT_VERSION As String = "[^"\n]+")'
    code, count = re.subn(
        version_pattern,
        lambda m: m[1] + '\nPrivate Const RELEASE_DATE_ISO As String = "' + iso + '"',
        code,
    )
    if count != 1:
        raise RuntimeError("Expected exactly one version declaration")
    code += "\nPublic Function AboutBuildDate() As String\n    AboutBuildDate = RELEASE_DATE_ISO\nEnd Function\n"
    assignment = "Me.VersionNumber.Caption = InterfaceVersionUpdate.InstalledVersion"
    if len(re.findall(re.escape(assignment), about_code, flags=re.IGNORECASE)) != 1:
        raise RuntimeError("Unexpected About version assignment")
    about_code = re.sub(
        re.escape(assignment),
        lambda m: m[0]
        + "\n   Me.Label2.Caption = InterfaceVersionUpdate.AboutBuildDate",
        about_code,
        flags=re.IGNORECASE,
    )
    return code, about_code


def main():
    import win32com.client

    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("target", type=Path)
    parser.add_argument("--managed-updates", action="store_true")
    parser.add_argument("--stamp-date", action="store_true")
    parser.add_argument("--build-date", default=date.today().isoformat())
    args = parser.parse_args()
    if args.stamp_date and not args.managed_updates:
        parser.error("--stamp-date requires --managed-updates")
    source, target = args.source.resolve(), args.target.resolve()
    version = "6.02.007" if args.managed_updates else "6.02.006"
    if args.stamp_date:
        version = "6.02.008"
    if target.exists() or source == target:
        raise RuntimeError("Use a new isolated output file")
    if hashlib.sha256(source.read_bytes()).hexdigest().upper() != BASELINE:
        raise RuntimeError("Source is not the approved baseline")
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, target)
    excel = win32com.client.DispatchEx("Excel.Application")
    book = None
    try:
        excel.AutomationSecurity = 3
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        book = excel.Workbooks.Open(str(target), UpdateLinks=0, ReadOnly=False)
        project = book.VBProject
        if project.Protection != 0:
            raise RuntimeError("VBA project is locked; no protection was changed")
        module = project.VBComponents.Item("InterfaceVersionUpdate").CodeModule
        code = module.Lines(1, module.CountOfLines)
        matches = list(
            re.finditer(
                r'^Private Const CURRENT_VERSION As String = "6\.02\.005"\s*$',
                code,
                re.MULTILINE,
            )
        )
        if len(matches) != 1:
            raise RuntimeError("Expected exactly one baseline version declaration")
        line = code[: matches[0].start()].count("\n") + 1
        module.ReplaceLine(
            line, f'Private Const CURRENT_VERSION As String = "{version}"'
        )
        if args.managed_updates:
            updated = managed_bridge(
                module.Lines(1, module.CountOfLines).replace("\r\n", "\n")
            )
            module.DeleteLines(1, module.CountOfLines)
            module.AddFromString(updated)
        about = project.VBComponents.Item("dlgAbout").CodeModule
        about_code = about.Lines(1, about.CountOfLines)
        header = "Private Sub UserForm_Initialize()"
        if about_code.count(header) != 1:
            raise RuntimeError("Unexpected About form code")
        line = about_code[: about_code.index(header)].count("\n") + 2
        about.InsertLines(line, '    Me.Caption = "NADABAS - SIMULATED UPGRADE TEST"')
        if args.stamp_date:
            new_code, new_about = stamp_build_date(
                module.Lines(1, module.CountOfLines),
                about.Lines(1, about.CountOfLines),
                args.build_date,
            )
            module.DeleteLines(1, module.CountOfLines)
            module.AddFromString(new_code)
            about.DeleteLines(1, about.CountOfLines)
            about.AddFromString(new_about)
        book.Save()
        evidence = {
            "testOnly": True,
            "publishedRelease": False,
            "fromVersion": "6.02.005",
            "toVersion": version,
            "sourceSha256": BASELINE,
            "candidate": str(target),
            "changes": ["CURRENT_VERSION", "About window test label"],
            "managedUpdateBridge": args.managed_updates,
            "buildDate": args.build_date if args.stamp_date else None,
            "signing": "must-sign-and-verify-before-packaging",
        }
        target.with_suffix(".json").write_text(json.dumps(evidence, indent=2))
    finally:
        if book is not None:
            book.Close(False)
        excel.Quit()
    if hashlib.sha256(source.read_bytes()).hexdigest().upper() != BASELINE:
        raise RuntimeError("Source changed unexpectedly")
    print(target)


if __name__ == "__main__":
    main()
