import csv
import unittest
from pathlib import Path

REPOSITORY_ROOT = Path(__file__).resolve().parents[1]


def read_vba(*parts: str) -> str:
    return REPOSITORY_ROOT.joinpath(*parts).read_text(encoding="utf-8-sig")


def procedure(source: str, declaration: str, terminator: str) -> str:
    start = source.index(declaration)
    end = source.index(terminator, start) + len(terminator)
    return source[start:end]


class VbaWorkbookMessageTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.open_workbooks = read_vba("vba", "modules", "cmdOpenWorkbooks.bas")
        cls.manage_workbooks = read_vba("vba", "forms", "frmManageWorkBooks.frm")

        resource_path = REPOSITORY_ROOT / "vba" / "resources" / "office-ui.csv"
        with resource_path.open(encoding="utf-8-sig", newline="") as handle:
            cls.resources = list(csv.DictReader(handle))

    def test_missing_workbook_explains_that_the_file_may_have_changed(self):
        missing_label = self.open_workbooks.index("\nnofile:\n") + 1
        missing_file = self.open_workbooks[
            missing_label : self.open_workbooks.index("OpenError:", missing_label)
        ]

        self.assertIn('GetMsg("M117")', missing_file)
        self.assertIn("WBtoOpen.path", missing_file)
        self.assertIn("WBtoOpen.WorkbookName", missing_file)
        self.assertIn('GetMsg("M117B")', missing_file)
        self.assertIn("vbExclamation", missing_file)

    def test_replace_warns_before_file_selection_and_defaults_to_no(self):
        replace = procedure(
            self.manage_workbooks, "Private Sub cmdReplaceFile_Click()", "End Sub"
        )

        warning = replace.index('GetMsg1("M215A", MWB.WorkbookName)')
        file_selection = replace.index("NewFileName = FileOpenDialog")
        self.assertLess(warning, file_selection)
        self.assertIn('GetMsg("M215B")', replace)
        self.assertIn('GetMsg("M215C")', replace)
        self.assertIn("vbDefaultButton2", replace)

    def test_replace_reports_that_data_must_be_loaded_again_after_success(self):
        replace = procedure(
            self.manage_workbooks, "Private Sub cmdReplaceFile_Click()", "End Sub"
        )

        self.assertLess(
            replace.index("ReplaceFile MWB.WorkbookName, NewName, newpath"),
            replace.index("ReplacementCompleted = True"),
        )
        self.assertLess(
            replace.index("ReplacementCompleted = True"),
            replace.index('MsgBox GetMsg("M216")'),
        )

    def test_new_messages_have_all_supported_translations(self):
        messages = {
            row["key"]: row
            for row in self.resources
            if row["sheet"] == "Messages"
        }

        self.assertEqual(
            set(messages), {"M117", "M117B", "M215A", "M215B", "M215C", "M216"}
        )
        self.assertIn("moved or renamed", messages["M117B"]["english"])
        self.assertIn("déplacé ou renommé", messages["M117B"]["french"])
        self.assertIn("movido ou renomeado", messages["M117B"]["portuguese"])
        self.assertIn("all data loaded", messages["M215B"]["english"])
        self.assertIn("Load the data", messages["M216"]["english"])
        for message in messages.values():
            for language in ("english", "french", "portuguese"):
                self.assertTrue(message[language].strip())


if __name__ == "__main__":
    unittest.main()
