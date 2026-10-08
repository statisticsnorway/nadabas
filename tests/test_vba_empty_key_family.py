"""Regression guards for the empty-only schema editing boundary."""

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def source(name):
    return (ROOT / "vba" / name).read_text(encoding="utf-8")


def test_policy_counts_rows_not_non_null_values_and_fails_closed():
    policy = source("modules/KeyFamilyEditPolicy.bas")
    assert "SELECT TOP 1 1 AS FoundRow FROM [" in policy
    assert "COUNT(Value)" not in policy
    assert "TableIsEmptyForStructure = Records.EOF" in policy
    assert "Unavailable:\n    TableIsEmptyForStructure = False" in policy
    assert "If Not isAdministrator Then Exit Function" in policy


def test_live_guard_precedes_schema_changes_and_commit():
    migration = source("modules/KeyFamilyMigration.bas").split(
        "Private Function BuildDraftDimensions"
    )[0]
    markers = [
        "CanEditEmptyKeyFamily(",
        "CurrentDB.DBCnn.BeginTrans",
        "CurrentDB.DBCnn, KeyFamilyName, ErrorMessage,",
        "CreateNewKeyFam(TemporaryTableName",
        "KeyFamilyName, BackupTableName, ErrorMessage)",
        "CurrentDB.DBCnn, BackupTableName, ErrorMessage)",
        "TemporaryTableName, KeyFamilyName, ErrorMessage)",
        "CurrentDB.DBCnn.CommitTrans",
    ]
    offsets = [migration.index(marker) for marker in markers]
    assert offsets == sorted(offsets)
    assert "CurrentDB.DBCnn.RollbackTrans" in migration
    assert migration.index("adModeShareExclusive") < migration.index(
        "CurrentDB.DBCnn.BeginTrans"
    )
    assert "CurrentDB.DBConnectionString, PreviousConnectionMode" in migration
    assert "WITH (TABLOCKX, HOLDLOCK)" in source("modules/KeyFamilyEditPolicy.bas")
    assert "DROP TABLE" not in migration.upper()


def test_structure_entry_point_is_guarded():
    parent = source("forms/frmKeyFamily.frm")
    availability = parent.split("Private Sub UpdateStructureAvailability()")[1].split(
        "End Sub"
    )[0]
    assert 'Me.Controls("cmdEditStructure").Enabled = Allowed' in availability
    assert "CanEditEmptyKeyFamily" in availability
    assert "If Not ManageMode Then Exit Sub" in parent


def test_editor_stages_changes_without_importing_pr41_database_mutations():
    editor = source("forms/dlgEditKeyFamily.frm")
    assert "KeyFamilyMigration.ApplySchemaDraft" in editor
    assert "CanEditEmptyKeyFamily" in editor
    assert "karna-stats" in editor
    assert "60e43df42425552771a3ae37411fef8d85677d69" in editor
    for direct_mutation in [
        "DBCnn.Execute",
        "DROP TABLE",
        "dlgModifyKeyLength",
        "RenameKeyFamily(",
        "ReorderKeyFamilyDimensions(",
    ]:
        assert direct_mutation not in editor
