import unittest
from pathlib import Path

REPOSITORY_ROOT = Path(__file__).resolve().parents[1]


def read_vba(*parts: str) -> str:
    return (REPOSITORY_ROOT.joinpath(*parts)).read_text(encoding="utf-8-sig")


def procedure(source: str, declaration: str, terminator: str) -> str:
    start = source.index(declaration)
    end = source.index(terminator, start) + len(terminator)
    return source[start:end]


class VbaWorkbookIdentityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.schema = read_vba("vba", "modules", "DbCreateTables.bas")
        cls.migration = read_vba("vba", "modules", "WorkbookIdentitySchema.bas")
        cls.conversion = read_vba("vba", "modules", "cmdConvertDB.bas")
        cls.database_open = read_vba("vba", "modules", "cmdOpen_Close_Database.bas")
        cls.db_functions = read_vba("vba", "modules", "DbFunc.bas")
        cls.workbook_info = read_vba("vba", "classes", "clsWorkBookInfo.cls")
        cls.source_info = read_vba("vba", "modules", "SourceInfo.bas")
        cls.manage_workbooks = read_vba("vba", "forms", "frmManageWorkBooks.frm")
        cls.documentation = read_vba("vba", "forms", "frmDocumentation.frm")
        cls.permission = read_vba("vba", "classes", "clsPermission.cls")

    def test_catalog_generates_numeric_ids_and_allows_supported_name_length(self):
        self.assertIn(
            "Public Const WORKBOOK_NAME_MAX_LENGTH As Long = 255",
            self.migration,
        )
        catalog = procedure(
            self.schema, "Public Sub CreateTableWorkbookIdentities", "End Sub"
        )
        self.assertIn("INT IDENTITY(1,1)", catalog)
        self.assertIn("AUTOINCREMENT", catalog)
        self.assertIn("WORKBOOK_NAME_MAX_LENGTH", catalog)
        self.assertIn('IndexCol "WorkbookID"', catalog)
        self.assertIn('IndexCol "WorkbookName"', catalog)

    def test_dependent_tables_store_ids_and_keep_255_character_names(self):
        expected = (
            'AddLongCol "WorkbookID"',
            'AddLongCol "TargetWorkbookID"',
            'AddLongCol "SourceWorkbookID"',
            'AddStrCol "WorkBookName", WORKBOOK_NAME_MAX_LENGTH',
            'AddStrCol "Workbook", WORKBOOK_NAME_MAX_LENGTH',
            'AddStrCol "Workbookname", WORKBOOK_NAME_MAX_LENGTH',
            'AddStrCol "TargetWB", WORKBOOK_NAME_MAX_LENGTH',
            'AddStrCol "SourceWB", WORKBOOK_NAME_MAX_LENGTH',
        )
        for declaration in expected:
            self.assertIn(declaration, self.schema)

        self.assertNotIn('AddStrCol "TargetWB", 50', self.schema)
        self.assertNotIn('AddStrCol "SourceWB", 50', self.schema)

    def test_datalinks_unique_index_uses_ids_instead_of_long_names(self):
        create_links = procedure(
            self.schema, "Public Sub CreateTableDataLinks", "End Sub"
        )
        index = create_links[create_links.index('Prepareindex "PrimaryIndex"') :]

        self.assertIn('IndexCol "TargetWorkbookID"', index)
        self.assertIn('IndexCol "SourceWorkbookID"', index)
        self.assertNotIn('IndexCol "TargetWB"', index)
        self.assertNotIn('IndexCol "SourceWB"', index)

    def test_access_metadata_is_refreshed_during_schema_changes(self):
        index_exists = procedure(
            self.schema, "Private Function IndexExists", "End Function"
        )
        column_exists = procedure(
            self.schema, "Public Function DBColumnExists", "End Function"
        )

        self.assertIn("CurrentDB.DBCat.Tables.Refresh", index_exists)
        self.assertIn("CurrentDB.DBCat.Tables(TableName).Indexes.Refresh", index_exists)
        self.assertIn("CurrentDB.DBCat.Tables.Refresh", column_exists)
        self.assertIn(
            "CurrentDB.DBCat.Tables(TableName).Columns.Refresh", column_exists
        )

    def test_existing_databases_are_backfilled_once_without_deleting_rows(self):
        ensure = procedure(
            self.migration,
            "Public Function EnsureWorkbookIdentitySchema",
            "End Function",
        )
        self.assertIn("CreateTableSchemaMigrations", ensure)
        self.assertIn("BackfillWorkbookIdentityData", ensure)
        self.assertIn("RequireWorkbookIDColumns", ensure)
        self.assertIn("WidenLegacyWorkbookNameColumns", ensure)
        self.assertIn("CreateWorkbookIdentityIndexes", ensure)
        self.assertIn('"WorkbookIdentityV1"', self.migration)
        self.assertNotIn("DELETE FROM", ensure.upper())

        for mapping in (
            '"Workbooks", "WorkBookName", "WorkbookID"',
            '"Permissions", "WorkBookName", "WorkbookID"',
            '"Documents", "Workbook", "WorkbookID"',
            '"BatchList", "Workbookname", "WorkbookID"',
            '"Descriptions", "WorkbookName", "WorkbookID"',
            '"DescriptionDimensions", "WorkbookName", "WorkbookID"',
            '"DataLinks", "TargetWB", "TargetWorkbookID"',
            '"DataLinks", "SourceWB", "SourceWorkbookID"',
        ):
            self.assertIn(mapping, self.migration)

    def test_migration_runs_before_database_collections_are_loaded(self):
        init = procedure(
            self.database_open, "Private Function InitNewDb", "End Function"
        )
        self.assertLess(
            init.index("EnsureWorkbookIdentitySchema"),
            init.index("CurrentDB.LoadAdministrators"),
        )

    def test_runtime_write_and_rename_paths_use_numeric_ids(self):
        save = procedure(self.workbook_info, "Public Sub SaveInDB", "End Sub")
        self.assertIn("GetOrCreateWorkbookID", save)
        self.assertIn('where WorkbookID = " & CStr(WorkbookID)', save)
        self.assertIn('PutColumn "WorkbookID", WorkbookID', save)

        self.assertIn('PutColumn "TargetWorkbookID"', self.source_info)
        self.assertIn('PutColumn "SourceWorkbookID"', self.source_info)

        rename = procedure(
            self.manage_workbooks, "Private Sub ReplaceOrRename", "End Sub"
        )
        self.assertIn('where WorkbookID = " & CStr(WorkbookID)', rename)
        self.assertIn('where SourceWorkbookID = " & CStr(WorkbookID)', rename)
        self.assertIn('where TargetWorkbookID = " & CStr(WorkbookID)', rename)

    def test_reports_join_workbook_metadata_by_id(self):
        self.assertIn(
            "Workbooks.WorkbookID = Descriptions.WorkbookID",
            self.documentation,
        )
        self.assertIn(
            "Descriptions.WorkbookID = DescriptionDimensions.WorkbookID",
            self.documentation,
        )
        self.assertNotIn(
            "Workbooks.WorkBookName = Descriptions.WorkbookName",
            self.documentation,
        )

    def test_conversion_preserves_generated_ids(self):
        self.assertLess(
            self.conversion.index('CopyTable("WorkbookIdentities")'),
            self.conversion.index('CopyTable("Workbooks")'),
        )
        copy_ids = procedure(
            self.conversion,
            "Private Function CopyWorkbookIdentityTable",
            "End Function",
        )
        self.assertIn("SET IDENTITY_INSERT [WorkbookIdentities] ON", copy_ids)
        self.assertIn("SET IDENTITY_INSERT [WorkbookIdentities] OFF", copy_ids)
        self.assertIn("[WorkbookID], [WorkbookName]", copy_ids)

    def test_sql_literals_escape_apostrophes_for_both_engines(self):
        in_q = procedure(self.db_functions, "Public Function InQ", "End Function")
        self.assertEqual(in_q.count('Replace(s, "\'", "\'\'")'), 2)
        self.assertIn("Case Sqlexpress", in_q)
        self.assertIn("Case accdb, mdb", in_q)

    def test_runtime_rejects_names_beyond_supported_length(self):
        create_identity = procedure(
            self.migration,
            "Public Function GetOrCreateWorkbookID",
            "End Function",
        )
        rename_identity = procedure(
            self.migration,
            "Public Sub RenameWorkbookIdentity",
            "End Sub",
        )

        self.assertIn("Len(WorkbookName) > WORKBOOK_NAME_MAX_LENGTH", create_identity)
        self.assertIn("Len(NewName) > WORKBOOK_NAME_MAX_LENGTH", rename_identity)

    def test_permission_collection_keys_use_numeric_identity(self):
        permission_key = procedure(
            self.permission, "Public Property Get key", "End Property"
        )

        self.assertIn('CStr(WorkbookID) & "_" & user', permission_key)
        self.assertNotIn('WorkBookName & "_" & user', permission_key)


if __name__ == "__main__":
    unittest.main()
