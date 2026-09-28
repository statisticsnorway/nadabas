from pathlib import Path

ROOT = Path(__file__).parents[1]
DESIGN = (ROOT / "vba" / "modules" / "cmdDesign.bas").read_text(encoding="utf-8-sig")
MIGRATION = (ROOT / "vba" / "modules" / "KeyFamilyMigration.bas").read_text(
    encoding="utf-8-sig"
)
KEY_FAMILY_FORM = (ROOT / "vba" / "forms" / "frmKeyFamily.frm").read_text(
    encoding="utf-8-sig"
)
BUTTON_CLASS = (ROOT / "vba" / "classes" / "clsKeyFamilyMigrationButton.cls").read_text(
    encoding="utf-8-sig"
)
DB_CREATE = (ROOT / "vba" / "modules" / "DbCreateTables.bas").read_text(
    encoding="utf-8-sig"
)
OFFICE_UI = (ROOT / "vba" / "resources" / "office-ui.csv").read_text(
    encoding="utf-8-sig"
)


def procedure(source: str, declaration: str, terminator: str) -> str:
    start = source.index(declaration)
    end = source.index(terminator, start) + len(terminator)
    return source[start:end]


def test_get_column_names_has_range_based_entry_point():
    writer = procedure(
        DESIGN, "Public Function TryWriteKeyFamilyColumnNames", "End Function"
    )
    executable = "\n".join(
        line for line in writer.splitlines() if not line.lstrip().startswith("'")
    )

    assert "TargetCell As Range" in writer
    assert "KeyFamilyName As String" in writer
    assert "ActiveCell" not in executable
    assert "Selection" not in executable
    assert "Set ResultRange = TargetCell.Resize" in writer


def test_dimension_rename_validates_all_definitions_before_applying_changes():
    rename = procedure(
        MIGRATION,
        "Public Function RenameDimensionInWorkbookDefinitions",
        "End Function",
    )

    validation = rename.index("' Validate every definition first")
    application = rename.index("If ApplyChange Then")
    assert validation < application
    assert "ValidateDimensionRename" in rename
    assert "ApplyDimensionRename" in rename


def test_dimension_rename_only_changes_the_field_name_column():
    apply_rename = procedure(MIGRATION, "Private Sub ApplyDimensionRename", "End Sub")

    assert "DBDefRange.Cells(i, 1).value = NewDimensionName" in apply_rename
    assert "Cells(i, 2).value" not in apply_rename


def test_key_family_definitions_are_discovered_through_dblinks():
    discovery = procedure(
        MIGRATION, "Public Function GetKeyFamilyDBDefinitionNames", "End Function"
    )

    assert "GetDBLinksRange(awb)" in discovery
    assert "DBLinksRange.Cells(k, 3).value" in discovery
    assert "DBDefinitionUsesKeyFamily" in discovery


def test_schema_change_buttons_are_in_the_existing_key_family_form():
    for control_name, caption, change_type in (
        ("cmdAddDimension", "Add dimension", "ADD"),
        ("cmdRemoveDimension", "Remove dimension", "REMOVE"),
        ("cmdRenameDimension", "Rename dimension", "RENAME"),
        ("cmdChangeDimensionLength", "Change length", "LENGTH"),
        ("cmdApplySchemaChanges", "Apply DB changes", "APPLY"),
        ("cmdPrepareWorkbookMigration", "Prepare workbook", "PREPAREWORKBOOK"),
        ("cmdApplyWorkbookMigration", "Activate DBDef draft", "APPLYWORKBOOK"),
    ):
        assert control_name in KEY_FAMILY_FORM
        assert f'"{change_type}"' in KEY_FAMILY_FORM
        assert f"Forms,frmKeyFamily,{control_name},{caption}" in OFFICE_UI

    assert "Private WithEvents ActionButton As MSForms.CommandButton" in BUTTON_CLASS
    assert "Owner.SchemaChangeButtonClick ChangeType" in BUTTON_CLASS
    assert "KeyFamilyMigration.ApplySchemaDraft(" in KEY_FAMILY_FORM


def test_schema_buttons_edit_a_draft_not_the_database():
    for declaration in (
        "Private Sub AddDimensionDraft",
        "Private Sub RemoveDimensionDraft",
        "Private Sub RenameDimensionDraft",
        "Private Sub ChangeDimensionLengthDraft",
    ):
        draft_editor = procedure(KEY_FAMILY_FORM, declaration, "End Sub")
        assert "DbExecute" not in draft_editor
        assert "DropTable" not in draft_editor

    review = procedure(MIGRATION, "Public Sub ReviewSchemaDraft", "End Sub")
    assert "SELECT COUNT(*) AS RowCount FROM" in review
    assert "WorkbooksUsingKeyFamily.GetWbForKey" in review
    assert "The database has not been changed." in review


def test_schema_button_dispatch_handles_unknown_actions():
    dispatch = procedure(
        KEY_FAMILY_FORM, "Public Sub SchemaChangeButtonClick", "End Sub"
    )

    assert "Select Case UCase$(ChangeType)" in dispatch
    assert "Case Else" in dispatch
    assert "Unsupported schema change action" in dispatch


def test_menu_preview_is_non_destructive():
    preview = procedure(MIGRATION, "Public Sub PreviewSchemaChange", "End Sub")

    assert 'Case "ADD"' in preview
    assert 'Case "REMOVE"' in preview
    assert 'Case "RENAME"' in preview
    assert "InputBox" not in preview


def test_apply_schema_requires_confirmation_before_database_changes():
    apply_schema = procedure(
        MIGRATION, "Public Function ApplySchemaDraft", "End Function"
    )

    confirmation = apply_schema.index('"Apply key-family schema"')
    create_replacement = apply_schema.index("CreateNewKeyFam(")
    rename_original = apply_schema.index("KeyFamilyName, BackupTableName, ErrorMessage")
    promote_replacement = apply_schema.index(
        "TemporaryTableName, KeyFamilyName, ErrorMessage"
    )

    assert confirmation < create_replacement < rename_original < promote_replacement
    assert "vbDefaultButton2" in apply_schema
    assert "The rebuilt key family will be empty" in apply_schema


def test_apply_schema_verifies_replacement_before_renaming_original():
    apply_schema = procedure(
        MIGRATION, "Public Function ApplySchemaDraft", "End Function"
    )
    verifier = procedure(
        MIGRATION, "Private Function VerifyReplacementTable", "End Function"
    )

    assert apply_schema.index("VerifyReplacementTable(TemporaryTableName") < (
        apply_schema.index("KeyFamilyName, BackupTableName, ErrorMessage")
    )
    assert "DBTableExists(TableName)" in verifier
    assert "DBColumnExists(TableName, field.name)" in verifier
    assert 'DBIndexExists(TableName, "PrimaryIndex")' in verifier


def test_apply_schema_keeps_recovery_table_and_has_reverse_order_rollback():
    apply_schema = procedure(
        MIGRATION, "Public Function ApplySchemaDraft", "End Function"
    )

    assert "DropTable" not in apply_schema
    rollback_start = apply_schema.index("' Roll back in reverse order")
    rollback = apply_schema[rollback_start:]
    move_replacement_aside = rollback.index(
        "KeyFamilyName, TemporaryTableName, RollbackMessage"
    )
    restore_original = rollback.index("BackupTableName, KeyFamilyName, RollbackMessage")
    assert move_replacement_aside < restore_original
    assert '"No table was deleted.' in apply_schema


def test_table_rename_supports_access_and_schema_qualified_sql_server():
    rename_table = procedure(
        DB_CREATE,
        "Public Function RenameTableForKeyFamilyMigration",
        "End Function",
    )

    assert "Case Sqlexpress" in rename_table
    assert 'SqlString("dbo." & OldTableName)' in rename_table
    assert "SqlString(NewTableName)" in rename_table
    assert "\", N'OBJECT'\"" in rename_table
    assert "Case accdb, mdb" in rename_table
    assert "CurrentDB.DBCat.Tables(OldTableName).name = NewTableName" in rename_table
    assert (
        "DBTableExists(OldTableName) Or Not DBTableExists(NewTableName)" in rename_table
    )


def test_workbook_migration_prepares_copies_without_switching_dblinks():
    prepare = procedure(
        MIGRATION, "Public Function PrepareActiveWorkbookMigration", "End Function"
    )

    assert "GetKeyFamilyDBDefinitionNames" in prepare
    assert "UniqueMigrationDefinitionName" in prepare
    assert "awb.Names.Add name:=ProposedName" in prepare
    assert "DBLinksRange.Cells" not in prepare
    assert "original definitions will not be" in prepare


def test_workbook_draft_is_validated_and_rolls_back_dblinks_on_failure():
    apply_workbook = procedure(
        MIGRATION, "Public Function ApplyActiveWorkbookMigration", "End Function"
    )

    preflight = apply_workbook.index("ValidateWorkbookMigrationDrafts(")
    confirmation = apply_workbook.index('"Activate DBDef draft"')
    switch = apply_workbook.index("DBLinksRange.Cells(k, 3).value =")
    validation = apply_workbook.index("TestDefinitions(awb, True)")
    rollback = apply_workbook.index("RestoreWorkbookDBLinks")
    assert preflight < confirmation < switch < validation < rollback
    assert "vbDefaultButton2" in apply_workbook
    assert "original DBDef ranges were retained" in apply_workbook


def test_workbook_preflight_targets_the_first_draft_error_before_dblinks():
    preflight = procedure(
        MIGRATION,
        "Private Function ValidateWorkbookMigrationDrafts",
        "End Function",
    )
    apply_workbook = procedure(
        MIGRATION, "Public Function ApplyActiveWorkbookMigration", "End Function"
    )

    assert "CountDBDefFieldOccurrences" in preflight
    assert '" is missing a mapping for dimension "' in preflight
    assert '" still contains "' in preflight
    assert "Set FirstErrorCell = DefinitionRange.Cells(j, 2)" in preflight
    assert "Application.Goto FirstDraftError, True" in apply_workbook
    before_confirmation = apply_workbook[
        : apply_workbook.index('"Activate DBDef draft"')
    ]
    assert "DBLinksRange.Cells(k, 3).value =" not in before_confirmation


def test_key_family_form_uses_staged_three_column_layout():
    layout = procedure(
        KEY_FAMILY_FORM, "Private Sub ConfigureKeyFamilyLayout", "End Sub"
    )

    for section in (
        "lblKeyFamilySection",
        "lblDimensionsSection",
        "lblWorkbooksSection",
        "lblMaintenanceSection",
    ):
        assert section in layout
    assert "cmdModify.Visible = False" in KEY_FAMILY_FORM
    assert "cmdDelete.BackColor" in layout
    assert ".Merge" not in layout
