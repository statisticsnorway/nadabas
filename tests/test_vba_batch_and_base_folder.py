import unittest
from pathlib import Path

REPOSITORY_ROOT = Path(__file__).resolve().parents[1]


def read_vba(*parts: str) -> str:
    return REPOSITORY_ROOT.joinpath(*parts).read_text(encoding="utf-8-sig")


def procedure(source: str, declaration: str, terminator: str) -> str:
    start = source.index(declaration)
    end = source.index(terminator, start) + len(terminator)
    return source[start:end]


class CleanupNameDefinitionErrorTests(unittest.TestCase):
    def test_definition_error_stops_cleanup_before_using_missing_names(self):
        source = read_vba("vba", "modules", "cmdDesign.bas")
        cleanup = procedure(source, "Public Sub CleanupAreaNames", "End Sub")
        error_branch_start = cleanup.index("If Areanames Is Nothing Then")
        error_branch_end = cleanup.index("End If", error_branch_start)
        error_branch = cleanup[error_branch_start:error_branch_end]

        self.assertIn('MsgBox GetMsg("M049")', error_branch)
        self.assertIn("Exit Sub", error_branch)
        self.assertGreater(
            cleanup.index('Areanames.Add "DBSOURCEFILES"'), error_branch_end
        )


class BaseFolderRepairTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = read_vba("vba", "modules", "cmdBaseFolder.bas")

    def test_set_base_folder_repairs_blank_paths_before_reloading_workbooks(self):
        set_base = procedure(self.source, "Public Sub SetBaseFolder", "End Sub")

        self.assertLess(
            set_base.index("CurrentDB.LoadBasePath", set_base.index("saveBasePath")),
            set_base.index("RepairWorkbookPathsForBaseFolder"),
        )
        self.assertLess(
            set_base.index("RepairWorkbookPathsForBaseFolder"),
            set_base.index("CurrentDB.LoadWorkbookInfo"),
        )

    def test_workbook_loading_accepts_a_null_path_before_base_folder_repair(self):
        database_source = read_vba("vba", "classes", "clsDB.cls")
        load_workbooks = procedure(
            database_source, "Public Sub LoadWorkbookInfo", "End Sub"
        )

        self.assertIn('WBinfo.RelPath = CStr(GetColumnNull("Path"))', load_workbooks)
        self.assertNotIn('WBinfo.RelPath = GetColumn("Path")', load_workbooks)

    def test_all_workbook_paths_are_checked_and_repairs_are_stored_by_id(self):
        repair = procedure(
            self.source,
            "Public Function RepairWorkbookPathsForBaseFolder",
            "End Function",
        )

        self.assertIn(
            "SELECT [WorkbookID], [WorkBookName], [Path] FROM [Workbooks]", repair
        )
        self.assertNotIn("WHERE [Path]", repair)
        self.assertIn("If Not IsPortableWorkbookPath(StoredPath) Then", repair)
        self.assertIn("If TestBasePath(StoredPath) Then", repair)
        self.assertIn("DirectPaths(WorkbookName)", repair)
        self.assertIn("ScanFolderForMissingWorkbooks", repair)
        self.assertIn("DropBasePath", repair)
        self.assertIn('WHERE [WorkbookID] = " & CStr(RepairIDs(NameKey))', repair)
        self.assertIn("AmbiguousNames.Exists", repair)

    def test_portable_root_and_subfolder_paths_are_left_unchanged(self):
        portable = procedure(
            self.source,
            "Private Function IsPortableWorkbookPath",
            "End Function",
        )

        self.assertIn('If trimmedPath = "!" Then', portable)
        self.assertIn('Left$(trimmedPath, 2) = "!\\"', portable)
        self.assertIn('Left$(trimmedPath, 2) = "!/"', portable)
        self.assertIn("StrComp(StoredPath, trimmedPath, vbBinaryCompare)", portable)

    def test_folder_scan_recognizes_supported_excel_extensions(self):
        scan = procedure(
            self.source,
            "Private Sub ScanFolderForMissingWorkbooks",
            "End Sub",
        )

        self.assertIn('Case "xls", "xlsx", "xlsm", "xlsb"', scan)
        self.assertIn("GetFilesInFolder", scan)
        self.assertIn("GetSubFolders", scan)
        self.assertIn("FoundFolders(WorkbookName) = FolderPath", scan)


if __name__ == "__main__":
    unittest.main()
