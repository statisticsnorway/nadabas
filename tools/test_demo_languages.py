"""Verify VBA language lookup in the approved demo add-in with desktop Excel."""

import argparse
import hashlib
import json
from pathlib import Path

EXPECTED = "BDF847E456764A7D335708471927547B780E33706BEB2A5635337F1FCF3135ED"
ROOT = Path(__file__).resolve().parents[1]
ADDIN = (
    ROOT
    / "artifacts/beta-review/855DB9C8DE239533/candidate-0.7.2/NADABAS.BETA.6.02.006.xlam"
)
KEY = r"Software\VB and VBA Program Settings\NADABAS\Language"


def preference():
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, KEY) as k:
            return winreg.QueryValueEx(k, "Code")
    except FileNotFoundError:
        return None


def main():
    import winreg

    import win32com.client

    parser = argparse.ArgumentParser()
    parser.add_argument("--demo", action="store_true")
    args = parser.parse_args()
    addin = ADDIN
    expected_hash = EXPECTED
    report_name = "signed-excel-language-test.json"
    if args.demo:
        addin = ROOT / "artifacts/demo-language/sign-later/NADABAS.DEMO.6.02.006.xlam"
        expected_hash = (
            "8D87B0295547DB6F250F99EEC60781ACB1F54F8FF596559253DF55F1628799E0"
        )
        report_name = "signed-demo-excel-language-test.json"

    if hashlib.sha256(addin.read_bytes()).hexdigest().upper() != expected_hash:
        raise RuntimeError("Unexpected add-in bytes")
    # Never borrow or close an existing interactive Excel instance.
    import subprocess

    processes = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
        capture_output=True,
    )
    if processes.returncode or b'"EXCEL.EXE"' in processes.stdout.upper():
        raise RuntimeError("Close Excel before testing the saved language preference")
    before = preference()
    last = None
    rows = {}
    excel = win32com.client.DispatchEx("Excel.Application")
    book = None
    try:
        excel.AutomationSecurity = 2  # Use existing Office policy.
        excel.EnableEvents = False
        book = excel.Workbooks.Open(str(addin), UpdateLinks=0, ReadOnly=True)
        if not book.VBASigned:
            raise RuntimeError("A signed VBA project is required")
        prefix = "'" + addin.name + "'!"
        for code, expected in [
            (4, ["Langue", "À propos"]),
            (5, ["Idioma", "Sobre"]),
            (6, ["Bahasa", "Tentang"]),
        ]:
            last = str(code)
            excel.Run(prefix + "InterfaceRegistry.SetLanguageSetting", code)
            excel.Run(prefix + "MultiLang.LoadLanguage")
            actual = {
                key: excel.Run(prefix + "MultiLang.GetRibbon", key)
                for key in (
                    "btnLanguage",
                    "btnAbout",
                    "btnCheckNadabasVersion",
                    "grpDatabase",
                )
            }
            if [actual["btnLanguage"], actual["btnAbout"]] != expected:
                raise RuntimeError(f"Unexpected text for {code}: {actual}")
            rows[str(code)] = actual
    finally:
        try:
            if book is not None:
                book.Close(False)
        finally:
            try:
                excel.Quit()
            finally:
                current = preference()
                if current != before:
                    if current is None or str(current[0]) != last:
                        raise RuntimeError("Concurrent preference change detected")
                    with winreg.OpenKey(
                        winreg.HKEY_CURRENT_USER, KEY, 0, winreg.KEY_SET_VALUE
                    ) as k:
                        if before is None:
                            winreg.DeleteValue(k, "Code")
                        else:
                            winreg.SetValueEx(k, "Code", 0, before[1], before[0])
                if preference() != before:
                    raise RuntimeError("Language preference restoration failed")
    out = ROOT / "artifacts/demo-language"
    out.mkdir(parents=True, exist_ok=True)
    report = {
        "scope": "Signed Excel VBA setter, language loading and ribbon lookup; no visual click test",
        "sha256": expected_hash,
        "rows": rows,
        "originalPreferenceRestored": True,
    }
    (out / report_name).write_text(
        json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    print(json.dumps(rows))


if __name__ == "__main__":
    main()
