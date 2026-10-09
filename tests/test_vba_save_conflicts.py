"""Guard the authorization and lifetime boundaries of save-conflict answers.

These source checks supplement, rather than replace, Excel execution tests.
"""

import csv
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


class SaveConflictTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.save = (ROOT / "vba/modules/cmdSaveData_ExportData.bas").read_text()
        cls.batch = (ROOT / "vba/modules/BatchRun.bas").read_text()
        cls.form = (ROOT / "vba/forms/dlgPutLog.frm").read_text()

    def test_remembered_answers_do_not_bypass_administrator_check(self):
        conflict = self.save[self.save.index('If Otherbook <> "" Then') :]
        self.assertLess(
            conflict.index("If isAdministrator Then"),
            conflict.index("If ConflictAnswer = vbNo Then"),
        )
        self.assertLess(
            conflict.index("If isAdministrator Then"),
            conflict.index("If ConflictAnswer <> vbYes Then"),
        )

    def test_no_to_all_stops_before_any_write_in_the_conflicting_range(self):
        start = self.save.index("If ConflictAnswer = vbNo Then")
        branch = self.save[start : self.save.index("End If", start)]
        self.assertIn("CancelSave = True", branch)
        self.assertIn("GoTo quit", branch)
        self.assertLess(start, self.save.index("ValueChanged = SaveAll"))

    def test_new_standalone_save_resets_answer_but_read_only_test_does_not(self):
        self.assertIn(
            "If Not TestOnly And Not ConflictBatchActive Then ConflictAnswer = 0",
            self.save,
        )

    def test_batch_scope_covers_early_exit_repeat_and_error(self):
        wrapper = self.batch[
            self.batch.index("Public Sub DoBatchUpdate()") : self.batch.index(
                "Private Sub RunBatchUpdate()"
            )
        ]
        self.assertLess(
            wrapper.index("BeginSaveConflictBatch"), wrapper.index("RunBatchUpdate")
        )
        self.assertEqual(wrapper.count("EndSaveConflictBatch"), 2)
        self.assertIn("On Error GoTo ConflictScopeError", wrapper)
        self.assertLess(
            wrapper.rindex("EndSaveConflictBatch"), wrapper.index("err.Raise")
        )

    def test_only_explicit_all_choice_is_remembered(self):
        start = self.save.index("If dlgPutLog.ApplyToAll Then")
        end = self.save.index("If dlgPutLog.LogYesNo = False Then", start)
        self.assertIn("ConflictAnswer = vbYes", self.save[start:end])
        self.assertIn("ConflictAnswer = vbNo", self.save[start:end])
        for button in ("cmdYesAll", "cmdNoAll"):
            start = self.form.index(f"Private Sub {button}_Click()")
            handler = self.form[start : self.form.index("End Sub", start)]
            self.assertIn("ApplyToAll = True", handler)
        self.assertIn("cmdNo.Cancel = True", self.form)
        self.assertIn("cmdNo.Default = True", self.form)
        self.assertIn("Private Sub UserForm_QueryClose", self.form)

    def test_both_buttons_have_all_four_translations(self):
        with (ROOT / "vba/resources/office-ui.csv").open(encoding="utf-8") as stream:
            rows = [r for r in csv.DictReader(stream) if r["key"] == "dlgPutLog"]
        self.assertEqual({r["control"] for r in rows}, {"cmdYesAll", "cmdNoAll"})
        for row in rows:
            for language in ("english", "french", "portuguese", "indonesian"):
                self.assertTrue(row[language].strip())
