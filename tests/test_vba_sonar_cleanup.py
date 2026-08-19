import re
import unittest
from pathlib import Path

REPOSITORY_ROOT = Path(__file__).resolve().parents[1]
SONAR_CLEANUP_FILES = (
    "vba/forms/frmBatchRun.frm",
    "vba/forms/frmClassification.frm",
    "vba/forms/frmRegisterDoc.frm",
    "vba/modules/ScanTableDefTest.bas",
)


class VbaSonarCleanupTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.sources = {
            relative_path: (REPOSITORY_ROOT / relative_path).read_text(
                encoding="utf-8-sig"
            )
            for relative_path in SONAR_CLEANUP_FILES
        }

    def test_string_functions_use_typed_string_returns(self):
        untyped_string_function = re.compile(r"\b(?:Mid|Trim|UCase)\(")

        for relative_path, source in self.sources.items():
            with self.subTest(path=relative_path):
                self.assertIsNone(untyped_string_function.search(source))

    def test_if_conditions_avoid_empty_string_and_boolean_literal_comparisons(self):
        empty_string_comparison = re.compile(
            r'\b[A-Za-z_][A-Za-z0-9_.]*\s*(?:=|<>)\s*""(?!")'
        )
        boolean_literal_comparison = re.compile(r"\s=\s(?:True|False)\b")

        for relative_path, source in self.sources.items():
            for line_number, line in enumerate(source.splitlines(), start=1):
                if not re.match(r"^\s*(?:If|ElseIf)\b", line):
                    continue
                with self.subTest(path=relative_path, line=line_number):
                    self.assertIsNone(empty_string_comparison.search(line))
                    self.assertIsNone(boolean_literal_comparison.search(line))

    def test_unknown_date_format_is_rejected(self):
        source = self.sources["vba/modules/ScanTableDefTest.bas"]
        start = source.index("Private Function TestDateFormat")
        end = source.index("End Function", start)
        procedure = source[start:end]

        case_else = procedure.index("Case Else")
        exit_function = procedure.index("Exit Function", case_else)
        end_select = procedure.index("End Select", exit_function)
        success_result = procedure.index("TestDateFormat = True", end_select)

        self.assertLess(case_else, exit_function)
        self.assertLess(exit_function, end_select)
        self.assertLess(end_select, success_result)


if __name__ == "__main__":
    unittest.main()
