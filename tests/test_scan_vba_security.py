import json
import re
import unittest
from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
from pathlib import Path

from tools.scan_vba_security import (
    Approval,
    _fails,
    load_approvals,
    main,
    mask_vba_strings,
    scan_text,
    strip_vba_comment,
)


class ScanVbaSecurityTests(unittest.TestCase):
    def test_comments_do_not_create_findings(self):
        source = """\
' kill dropbox does save the commandline
vbOKOnly       'Not able to kill dropbox
Rem Shell "cmd.exe"
x = 1: Rem Kill target
"""

        self.assertEqual(scan_text("vba/modules/BatchRun.bas", source), [])

    def test_apostrophe_inside_string_is_preserved(self):
        line = 'message = "Bob\'s file" \' ordinary comment'

        self.assertEqual(strip_vba_comment(line), 'message = "Bob\'s file"')

    def test_security_words_inside_an_ordinary_string_are_masked(self):
        source = 'Caption = "Use Shell or SendKeys"'

        self.assertEqual(scan_text("vba/forms/Dialog.frm", source), [])
        self.assertNotIn("Shell", mask_vba_strings(source))

    def test_real_file_delete_is_an_error(self):
        findings = scan_text(
            "vba/modules/InstallMeAsAddin.bas",
            "    Kill AI.FullName ' replace the installed add-in",
        )

        self.assertEqual(len(findings), 1)
        self.assertEqual(findings[0].rule_id, "file-delete")
        self.assertEqual(findings[0].severity, "error")
        self.assertEqual(findings[0].code, "Kill AI.FullName")

    def test_inline_file_delete_is_an_error(self):
        findings = scan_text(
            "vba/modules/Example.bas", "If obsolete Then Kill targetPath"
        )

        self.assertEqual([finding.rule_id for finding in findings], ["file-delete"])

    def test_wmi_process_operations_are_detected(self):
        source = r"""\
Set objServices = GetObject("winmgmts:\\.\root\cimv2")
Set processes = objServices.ExecQuery("Select * from Win32_Process")
For Each process In processes
    process.Terminate
Next
objServices.Create("dropbox.exe"), Null, Null, processId
"""

        findings = scan_text("vba/modules/DropBoxINterface.bas", source)
        rule_ids = {finding.rule_id for finding in findings}

        self.assertEqual(
            rule_ids,
            {
                "wmi-connection",
                "wmi-query",
                "process-terminate",
                "wmi-process-create",
            },
        )

    def test_native_declaration_is_info_not_shell_execute(self):
        source = (
            'Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" '
            'Alias "ShellExecuteA" () As Long\n'
        )

        findings = scan_text("vba/classes/Example.cls", source)

        self.assertEqual([finding.rule_id for finding in findings], ["native-api"])
        self.assertEqual(findings[0].severity, "info")

    def test_shell_execute_call_is_a_warning(self):
        findings = scan_text(
            "vba/forms/frmOpenDocuments.frm",
            "result = ShellExecute(0, \"open\", file, vbNullString, path, 1)",
        )

        self.assertEqual([finding.rule_id for finding in findings], ["shell-execute"])
        self.assertEqual(findings[0].severity, "warning")

    def test_approval_is_limited_to_path_rule_and_line_shape(self):
        approval = Approval(
            path_glob="vba/modules/InstallMeAsAddin.bas",
            rule_id="file-delete",
            line_regex=re.compile(r"^\s*Kill\s+AI\.FullName$", re.IGNORECASE),
            reason="The installer replaces only the selected NADABAS add-in.",
        )

        approved = scan_text(
            "vba/modules/InstallMeAsAddin.bas",
            "Kill AI.FullName",
            approvals=(approval,),
        )[0]
        other_line = scan_text(
            "vba/modules/InstallMeAsAddin.bas",
            "Kill userSelectedPath",
            approvals=(approval,),
        )[0]
        other_path = scan_text(
            "vba/modules/Other.bas", "Kill AI.FullName", approvals=(approval,)
        )[0]

        self.assertTrue(approved.approved)
        self.assertFalse(other_line.approved)
        self.assertFalse(other_path.approved)

    def test_fail_threshold_ignores_approved_findings(self):
        info = scan_text("vba/classes/Example.cls", "Declare Function Beep Lib \"x\" ()")[0]
        warning = scan_text("vba/modules/Example.bas", "SendKeys \"x\"")[0]
        error = scan_text("vba/modules/Example.bas", "Kill target")[0]

        self.assertFalse(_fails([info, warning], "error"))
        self.assertTrue(_fails([info, warning], "warning"))
        self.assertTrue(_fails([error], "error"))
        self.assertFalse(_fails([error], "never"))

    def test_auto_open_is_reported_as_warning(self):
        findings = scan_text("vba/modules/AutoOpenClose.bas", "Public Sub Auto_Open()")

        self.assertEqual([finding.rule_id for finding in findings], ["auto-execution"])

    def test_invalid_root_returns_usage_error(self):
        stdout = StringIO()
        stderr = StringIO()

        with redirect_stdout(stdout), redirect_stderr(stderr):
            exit_code = main(["directory-that-does-not-exist"])

        self.assertEqual(exit_code, 2)
        self.assertIn("does not exist", stderr.getvalue())


class ApprovalFileTests(unittest.TestCase):
    def test_approval_file_requires_a_reason(self):
        path = Path("approval-without-reason.json")
        payload = {
            "approvals": [
                {
                    "path": "vba/modules/Example.bas",
                    "rule_id": "file-delete",
                    "line_regex": "Kill target",
                    "reason": "",
                }
            ]
        }

        # Mock a Path read without creating files in the test environment.
        class MemoryPath:
            def read_text(self, encoding):
                self.encoding = encoding
                return json.dumps(payload)

            def __str__(self):
                return str(path)

        with self.assertRaisesRegex(ValueError, "non-empty reason"):
            load_approvals(MemoryPath())

    def test_approval_file_rejects_wildcard_paths(self):
        payload = {
            "approvals": [
                {
                    "path": "vba/**/*.bas",
                    "rule_id": "file-delete",
                    "line_regex": "Kill target",
                    "reason": "Too broad",
                }
            ]
        }

        class MemoryPath:
            def read_text(self, encoding):
                return json.dumps(payload)

            def __str__(self):
                return "broad-approval.json"

        with self.assertRaisesRegex(ValueError, "exact file"):
            load_approvals(MemoryPath())


if __name__ == "__main__":
    unittest.main()
