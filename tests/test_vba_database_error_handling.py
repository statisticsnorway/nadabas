import unittest
from pathlib import Path

REPOSITORY_ROOT = Path(__file__).resolve().parents[1]


def vba_procedure(source: str, declaration: str, terminator: str) -> str:
    start = source.index(declaration)
    end = source.index(terminator, start) + len(terminator)
    return source[start:end]


class VbaDatabaseErrorHandlingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.db_interface = (
            REPOSITORY_ROOT / "vba" / "modules" / "DBInterface.bas"
        ).read_text(encoding="ascii")
        cls.open_database = (
            REPOSITORY_ROOT / "vba" / "modules" / "cmdOpen_Close_Database.bas"
        ).read_text(encoding="ascii")
        cls.batch_run = (
            REPOSITORY_ROOT / "vba" / "modules" / "BatchRun.bas"
        ).read_text(encoding="ascii")

    def test_create_cursor_retains_the_original_ado_error(self):
        procedure = vba_procedure(
            self.db_interface, "Public Function CreateCursor", "End Function"
        )

        self.assertIn("LastCursorSql = ssql", procedure)
        self.assertIn("LastCursorErrorNumber = errorNumber", procedure)
        self.assertIn("LastCursorErrorDescription = errorDescription", procedure)
        self.assertIn("Set rs = Nothing", procedure)

    def test_cursor_add_new_checks_that_the_recordset_is_open(self):
        procedure = vba_procedure(
            self.db_interface, "Public Sub CursorAddNew", "End Sub"
        )

        guard = 'If Not CursorIsOpen Then RaiseCursorNotOpen "CursorAddNew"'
        self.assertLess(procedure.index(guard), procedure.index("rs.AddNew"))

    def test_create_snapshot_preserves_database_and_provider_errors(self):
        procedure = vba_procedure(
            self.db_interface, "Public Sub CreateSnapshot", "End Sub"
        )

        self.assertIn("errorNumber = err.Number", procedure)
        self.assertIn("errorDescription = err.Description", procedure)
        self.assertIn("For Each providerError In CurrentDB.DBCnn.Errors", procedure)
        self.assertIn("providerError.Description", procedure)
        self.assertIn("Set rs = Nothing", procedure)
        self.assertIn("AddToSQLLog message", procedure)
        self.assertIn(
            'err.Raise errorNumber, "DBInterface.CreateSnapshot", message', procedure
        )
        self.assertNotIn("MsgBox ssql", procedure)

    def test_batch_error_replaces_workbook_placeholder(self):
        procedure = vba_procedure(
            self.batch_run, "Private Sub HandleBatchError", "End Sub"
        )

        self.assertIn('GetMsg1("M014", SI.WorkbookName)', procedure)
        self.assertNotIn('GetMsg("M014")', procedure)

    def test_open_database_handles_and_cleans_up_unexpected_errors(self):
        procedure = vba_procedure(
            self.open_database, "Public Sub OpenDatabase", "End Sub"
        )

        self.assertIn("On Error GoTo OpenError", procedure)
        self.assertIn("OpenError:", procedure)
        self.assertIn("CloseDbAll", procedure)
        self.assertIn("Set CurrentDB.DBCnn = Nothing", procedure)
        self.assertIn("SetNadabasIsSleeping (True)", procedure)
        self.assertIn('MsgBox errorMessage, vbCritical, "NADABAS"', procedure)

        user_message = procedure[
            procedure.index('errorMessage = "NADABAS could not open') : procedure.index(
                "MsgBox errorMessage"
            )
        ]
        self.assertNotIn("errorNumber", user_message)
        self.assertNotIn("errorDescription", user_message)
        self.assertIn("select Enable Content", user_message)
        self.assertIn("Technical details are available", user_message)


if __name__ == "__main__":
    unittest.main()
