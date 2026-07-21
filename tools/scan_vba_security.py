"""Static, review-oriented security scanner for exported VBA source files.

The scanner is intentionally conservative: it identifies operations that need
human review, not whether the surrounding business logic is malicious.  Rules
have severities so CI can fail on errors while still reporting warnings and
informational findings.
"""

from __future__ import annotations

import argparse
import fnmatch
import json
import re
import sys
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable, Sequence


VBA_EXTENSIONS = {".bas", ".cls", ".frm"}
SEVERITY_RANK = {"info": 1, "warning": 2, "error": 3}


@dataclass(frozen=True)
class Rule:
    """One security-relevant VBA pattern."""

    rule_id: str
    severity: str
    message: str
    pattern: re.Pattern[str]
    required_context: str | None = None
    inspect_strings: bool = False


@dataclass(frozen=True)
class Approval:
    """A narrow, documented exception for one rule and one line shape."""

    path_glob: str
    rule_id: str
    line_regex: re.Pattern[str]
    reason: str


@dataclass(frozen=True)
class Finding:
    """A rule match in one source line."""

    path: str
    line: int
    rule_id: str
    severity: str
    message: str
    code: str
    approved: bool = False
    approval_reason: str | None = None


RULES = (
    Rule(
        "process-shell",
        "error",
        "Starts an external command through VBA Shell.",
        re.compile(
            r"^(?!\s*(?:(?:Public|Private|Friend)\s+)?(?:Sub|Function)\s+Shell\b)"
            r".*\bShell\b\s*(?:\(|\s)",
            re.IGNORECASE,
        ),
    ),
    Rule(
        "wscript-shell",
        "error",
        "Creates WScript.Shell, which can execute commands and modify the registry.",
        re.compile(r"\bCreateObject\s*\(\s*\"WScript\.Shell\"", re.IGNORECASE),
        inspect_strings=True,
    ),
    Rule(
        "command-interpreter",
        "error",
        "References a command or script interpreter.",
        re.compile(
            r"\b(?:powershell(?:\.exe)?|pwsh(?:\.exe)?|cmd\.exe|wscript\.exe|cscript\.exe)\b",
            re.IGNORECASE,
        ),
        inspect_strings=True,
    ),
    Rule(
        "process-terminate",
        "error",
        "Terminates a process through an object API.",
        re.compile(r"\.\s*Terminate\b", re.IGNORECASE),
    ),
    Rule(
        "url-download",
        "error",
        "Downloads content through URLDownloadToFile.",
        re.compile(
            r"^(?!\s*(?:(?:Private|Public)\s+)?Declare\b)"
            r".*\bURLDownloadToFile(?:A|W)?\b",
            re.IGNORECASE,
        ),
    ),
    Rule(
        "file-delete",
        "error",
        "Deletes a file.",
        re.compile(
            r"(?:^|:\s*|\bThen\s+)\s*(?:Kill|RmDir)\s+|\.\s*DeleteFile\b",
            re.IGNORECASE,
        ),
    ),
    Rule(
        "send-keys",
        "warning",
        "Sends synthetic keyboard input.",
        re.compile(r"\bSendKeys\b", re.IGNORECASE),
    ),
    Rule(
        "shell-execute",
        "warning",
        "Opens a file or URL through the operating-system shell.",
        re.compile(
            r"^(?!\s*(?:(?:Private|Public)\s+)?Declare\b).*\bShellExecute(?:A|W)?\b",
            re.IGNORECASE,
        ),
    ),
    Rule(
        "wmi-connection",
        "warning",
        "Connects to Windows Management Instrumentation (WMI).",
        re.compile(r"\b(?:GetObject|CreateObject)\s*\(\s*\"winmgmts:", re.IGNORECASE),
        inspect_strings=True,
    ),
    Rule(
        "wmi-process-create",
        "warning",
        "Creates a process through WMI.",
        re.compile(r"\.\s*Create\s*\(", re.IGNORECASE),
        required_context="wmi-process",
    ),
    Rule(
        "filesystem-object",
        "warning",
        "Creates or references Scripting.FileSystemObject.",
        re.compile(r"\bScripting\.FileSystemObject\b", re.IGNORECASE),
        inspect_strings=True,
    ),
    Rule(
        "http-client",
        "warning",
        "Creates or references an HTTP client.",
        re.compile(r"\b(?:WinHttp\.WinHttpRequest|MSXML2\.XMLHTTP)\b", re.IGNORECASE),
        inspect_strings=True,
    ),
    Rule(
        "adodb-stream",
        "warning",
        "Uses ADODB.Stream, which can write downloaded or binary content.",
        re.compile(r"\bADODB\.Stream\b", re.IGNORECASE),
        inspect_strings=True,
    ),
    Rule(
        "file-copy",
        "warning",
        "Copies a file.",
        re.compile(r"^\s*FileCopy\b", re.IGNORECASE),
    ),
    Rule(
        "file-write",
        "warning",
        "Opens a file for writing.",
        re.compile(r"^\s*Open\s+.+\s+For\s+(?:Output|Append|Binary)\b", re.IGNORECASE),
    ),
    Rule(
        "auto-execution",
        "warning",
        "Runs automatically when an Office document opens or closes.",
        re.compile(
            r"^\s*(?:(?:Public|Private|Friend)\s+)?Sub\s+"
            r"(?:Auto_Open|Auto_Close|Workbook_Open|Workbook_BeforeClose)\b",
            re.IGNORECASE,
        ),
    ),
    Rule(
        "native-api",
        "info",
        "Declares a native operating-system API.",
        re.compile(
            r"^\s*(?:(?:Private|Public)\s+)?Declare\s+(?:PtrSafe\s+)?(?:Function|Sub)\b",
            re.IGNORECASE,
        ),
    ),
    Rule(
        "environment-read",
        "info",
        "Reads an environment variable.",
        re.compile(r"\bEnviron\$?\s*\(", re.IGNORECASE),
    ),
    Rule(
        "wmi-query",
        "info",
        "Executes a WMI query.",
        re.compile(r"\.\s*ExecQuery\b", re.IGNORECASE),
    ),
    Rule(
        "directory-create",
        "info",
        "Creates a directory.",
        re.compile(r"^\s*MkDir\b", re.IGNORECASE),
    ),
)


def strip_vba_comment(line: str) -> str:
    """Remove a VBA apostrophe or ``Rem`` comment while preserving strings."""

    result: list[str] = []
    in_string = False
    index = 0
    while index < len(line):
        if not in_string and line[index : index + 3].casefold() == "rem":
            previous = "".join(result).rstrip()
            next_char = line[index + 3 : index + 4]
            if (not previous or previous.endswith(":")) and (
                not next_char or next_char.isspace()
            ):
                break
        char = line[index]
        if char == '"':
            result.append(char)
            if in_string and index + 1 < len(line) and line[index + 1] == '"':
                result.append('"')
                index += 2
                continue
            in_string = not in_string
        elif char == "'" and not in_string:
            break
        else:
            result.append(char)
        index += 1
    return "".join(result).rstrip()


def mask_vba_strings(code: str) -> str:
    """Mask string contents so identifiers in captions/text are not scanned."""

    result: list[str] = []
    in_string = False
    index = 0
    while index < len(code):
        char = code[index]
        if char == '"':
            result.append(char)
            if in_string and index + 1 < len(code) and code[index + 1] == '"':
                result.append(" ")
                index += 2
                continue
            in_string = not in_string
        else:
            result.append(" " if in_string else char)
        index += 1
    return "".join(result)


def _source_context(code_lines: Sequence[str]) -> set[str]:
    combined = "\n".join(code_lines).casefold()
    context: set[str] = set()
    if "winmgmts:" in combined and "win32_process" in combined:
        context.add("wmi-process")
    return context


def _approval_for(
    path: str, rule_id: str, code: str, approvals: Iterable[Approval]
) -> Approval | None:
    normalised = path.replace("\\", "/")
    for approval in approvals:
        if (
            approval.rule_id == rule_id
            and fnmatch.fnmatch(normalised, approval.path_glob)
            and approval.line_regex.search(code)
        ):
            return approval
    return None


def scan_text(
    path: str | Path,
    text: str,
    *,
    approvals: Iterable[Approval] = (),
    rules: Sequence[Rule] = RULES,
) -> list[Finding]:
    """Scan VBA text and return deterministic, structured findings."""

    display_path = str(path).replace("\\", "/")
    original_lines = text.splitlines()
    code_lines = [strip_vba_comment(line) for line in original_lines]
    masked_lines = [mask_vba_strings(line) for line in code_lines]
    context = _source_context(code_lines)
    findings: list[Finding] = []

    for line_number, (code, masked_code) in enumerate(
        zip(code_lines, masked_lines), start=1
    ):
        if not code.strip():
            continue
        for rule in rules:
            if rule.required_context and rule.required_context not in context:
                continue
            scan_target = code if rule.inspect_strings else masked_code
            if not rule.pattern.search(scan_target):
                continue
            approval = _approval_for(display_path, rule.rule_id, code, approvals)
            findings.append(
                Finding(
                    path=display_path,
                    line=line_number,
                    rule_id=rule.rule_id,
                    severity=rule.severity,
                    message=rule.message,
                    code=code.strip(),
                    approved=approval is not None,
                    approval_reason=approval.reason if approval else None,
                )
            )
    return findings


def _read_vba_text(path: Path) -> str:
    data = path.read_bytes()
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise UnicodeError(f"Could not decode {path}")


def scan_file(path: Path, *, approvals: Iterable[Approval] = ()) -> list[Finding]:
    """Read and scan one exported VBA component."""

    return scan_text(path, _read_vba_text(path), approvals=approvals)


def scan_tree(
    root: Path, *, approvals: Iterable[Approval] = ()
) -> tuple[list[Finding], int]:
    """Scan every supported VBA source below a directory."""

    paths = sorted(
        (
            path
            for path in root.rglob("*")
            if path.is_file() and path.suffix.lower() in VBA_EXTENSIONS
        ),
        key=lambda path: str(path).casefold(),
    )
    findings: list[Finding] = []
    for path in paths:
        findings.extend(scan_file(path, approvals=approvals))
    return findings, len(paths)


def load_approvals(path: Path | None) -> tuple[Approval, ...]:
    """Load narrow approvals from JSON; an approval must always include a reason."""

    if path is None:
        return ()
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"Could not read approval file {path}: {exc}") from exc

    if not isinstance(payload, dict):
        raise ValueError("Approval file must contain a JSON object.")
    raw_approvals = payload.get("approvals")
    if not isinstance(raw_approvals, list):
        raise ValueError("Approval file must contain an 'approvals' list.")

    approvals: list[Approval] = []
    valid_rule_ids = {rule.rule_id for rule in RULES}
    for index, item in enumerate(raw_approvals, start=1):
        if not isinstance(item, dict):
            raise ValueError(f"Approval {index} must be an object.")
        try:
            raw_path = item["path"]
            raw_rule_id = item["rule_id"]
            raw_line_regex = item["line_regex"]
            raw_reason = item["reason"]
        except KeyError as exc:
            raise ValueError(f"Approval {index} is missing {exc.args[0]!r}.") from exc
        if not all(
            isinstance(value, str)
            for value in (raw_path, raw_rule_id, raw_line_regex, raw_reason)
        ):
            raise ValueError(f"Approval {index} fields must all be strings.")
        path_glob = raw_path.replace("\\", "/").strip()
        rule_id = raw_rule_id.strip()
        line_regex = raw_line_regex
        reason = raw_reason.strip()
        if not path_glob or any(character in path_glob for character in "*?[]"):
            raise ValueError(
                f"Approval {index} path must name one exact file without wildcards."
            )
        if path_glob.startswith("/") or re.match(r"^[A-Za-z]:/", path_glob):
            raise ValueError(f"Approval {index} path must be repository-relative.")
        if ".." in Path(path_glob).parts:
            raise ValueError(f"Approval {index} path must not contain '..'.")
        if Path(path_glob).suffix.lower() not in VBA_EXTENSIONS:
            raise ValueError(f"Approval {index} path must name a VBA source file.")
        if rule_id not in valid_rule_ids:
            raise ValueError(f"Approval {index} uses unknown rule_id {rule_id!r}.")
        if not reason:
            raise ValueError(f"Approval {index} must include a non-empty reason.")
        try:
            compiled = re.compile(line_regex, re.IGNORECASE)
        except re.error as exc:
            raise ValueError(f"Approval {index} has invalid line_regex: {exc}") from exc
        approvals.append(Approval(path_glob, rule_id, compiled, reason))
    return tuple(approvals)


def _summary(findings: Iterable[Finding]) -> dict[str, int]:
    counts = {"error": 0, "warning": 0, "info": 0, "approved": 0}
    for finding in findings:
        if finding.approved:
            counts["approved"] += 1
        else:
            counts[finding.severity] += 1
    return counts


def _fails(findings: Iterable[Finding], fail_on: str) -> bool:
    if fail_on == "never":
        return False
    threshold = SEVERITY_RANK[fail_on]
    return any(
        not finding.approved and SEVERITY_RANK[finding.severity] >= threshold
        for finding in findings
    )


def _print_text(findings: Sequence[Finding], scanned_files: int, show_approved: bool) -> None:
    visible = [finding for finding in findings if show_approved or not finding.approved]
    for finding in visible:
        status = "APPROVED" if finding.approved else finding.severity.upper()
        print(
            f"{status:8} {finding.path}:{finding.line} [{finding.rule_id}] "
            f"{finding.message}\n         {finding.code}"
        )
        if finding.approved:
            print(f"         Reason: {finding.approval_reason}")

    counts = _summary(findings)
    print(
        "\nSummary: "
        f"files={scanned_files}, errors={counts['error']}, "
        f"warnings={counts['warning']}, info={counts['info']}, "
        f"approved={counts['approved']}"
    )


def _print_json(findings: Sequence[Finding], scanned_files: int) -> None:
    print(
        json.dumps(
            {
                "scanned_files": scanned_files,
                "summary": _summary(findings),
                "findings": [asdict(finding) for finding in findings],
            },
            indent=2,
            ensure_ascii=False,
        )
    )


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Scan exported VBA source for security-relevant operations."
    )
    parser.add_argument("root", nargs="?", type=Path, default=Path("vba"))
    parser.add_argument(
        "--fail-on",
        choices=("error", "warning", "info", "never"),
        default="error",
        help="Lowest unapproved severity that returns exit code 1 (default: error)",
    )
    parser.add_argument("--approvals", type=Path, help="JSON file with narrow approvals")
    parser.add_argument("--show-approved", action="store_true")
    parser.add_argument("--format", choices=("text", "json"), default="text")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    if not args.root.is_dir():
        print(f"ERROR: VBA source directory does not exist: {args.root}", file=sys.stderr)
        return 2
    try:
        approvals = load_approvals(args.approvals)
        findings, scanned_files = scan_tree(args.root, approvals=approvals)
    except (OSError, UnicodeError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    if args.format == "json":
        _print_json(findings, scanned_files)
    else:
        _print_text(findings, scanned_files, args.show_approved)
    return 1 if _fails(findings, args.fail_on) else 0


if __name__ == "__main__":
    raise SystemExit(main())
