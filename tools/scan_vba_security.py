from pathlib import Path
import sys

SUSPICIOUS_PATTERNS = [
    "Shell",
    "WScript.Shell",
    "Scripting.FileSystemObject",
    "URLDownloadToFile",
    "WinHttp.WinHttpRequest",
    "MSXML2.XMLHTTP",
    "ADODB.Stream",
    "Kill ",
    "Declare PtrSafe",
    "Declare Function",
    "Environ",
    "SendKeys",
    "Auto_Open",
    "Workbook_Open",
]

VBA_EXTENSIONS = {".bas", ".cls", ".frm"}


def scan_file(path: Path) -> list[str]:
    findings = []
    text = path.read_text(encoding="latin-1", errors="ignore")

    for lineno, line in enumerate(text.splitlines(), start=1):
        for pattern in SUSPICIOUS_PATTERNS:
            if pattern.lower() in line.lower():
                findings.append(f"{path}:{lineno}: found '{pattern}' -> {line.strip()}")

    return findings


def main() -> int:
    root = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("vba")

    findings = []
    for path in root.rglob("*"):
        if path.suffix.lower() in VBA_EXTENSIONS:
            findings.extend(scan_file(path))

    if findings:
        print("Potentially suspicious VBA patterns found:\n")
        for finding in findings:
            print(finding)
        return 1

    print("No suspicious VBA patterns found.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())