"""Execute repository VBA migration on new synthetic Access files only.

Uses an owned unsaved Excel instance and existing trusted VBA-project access.
Optional named test databases are copied before opening; originals are never
opened through ACE. Does not change Office policy or save an add-in.
"""

import argparse
import hashlib
import json
import re
import shutil
import tempfile
from collections import Counter
from pathlib import Path

IDENTITY_COLUMNS = {
    "workbooks": {"workbookid"},
    "permissions": {"workbookid"},
    "documents": {"workbookid"},
    "batchlist": {"workbookid"},
    "descriptions": {"workbookid"},
    "descriptiondimensions": {"workbookid"},
    "datalinks": {"targetworkbookid", "sourceworkbookid"},
}


def legacy_snapshot(path, schema=None):
    """Hash complete legacy rows in memory; never emit database values.

    COUNT(*) can use a stale Access index. Enumerate actual records instead.
    Exclude only the identity columns intentionally filled by this migration.
    Call exclusively on owned copies, never an original production database.
    """
    import win32com.client

    connection = win32com.client.Dispatch("ADODB.Connection")
    connection.Open(f"Provider=Microsoft.ACE.OLEDB.16.0;Mode=Read;Data Source={path}")
    try:
        if schema is None:
            catalog = win32com.client.Dispatch("ADOX.Catalog")
            catalog.ActiveConnection = connection
            schema = {}
            for table in catalog.Tables:
                if table.Type != "TABLE" or table.Name.lower() in {
                    "workbookidentities",
                    "nadabasschemamigrations",
                }:
                    continue
                excluded = IDENTITY_COLUMNS.get(table.Name.lower(), set())
                schema[table.Name] = [
                    column.Name
                    for column in table.Columns
                    if column.Name.lower() not in excluded
                ]
        snapshots = {}
        for table, columns in schema.items():
            bracket = lambda name: "[" + name.replace("]", "]]") + "]"
            query = (
                "SELECT " + ",".join(map(bracket, columns)) + " FROM " + bracket(table)
            )
            records = connection.Execute(query)[0]
            rows = Counter()
            while not records.EOF:
                for values in zip(*records.GetRows(4096)):
                    encoded = json.dumps(
                        values, ensure_ascii=False, default=str
                    ).encode("utf-8")
                    rows[hashlib.sha256(encoded).digest()] += 1
            records.Close()
            snapshots[table] = rows
        return schema, snapshots
    finally:
        connection.Close()


def procedures(source, names):
    chunks = []
    for name in names:
        match = re.search(
            rf"^[ \t]*(?:Public|Private) (Sub|Function) {name}\b.*?^[ \t]*End \1",
            source,
            re.M | re.S,
        )
        if match is None:
            raise ValueError(f"Missing VBA procedure: {name}")
        chunks.append(match.group())
    return "\n\n".join(chunks)


def verify_identity_references(path, schema):
    """Check both sides of every migrated name/ID association without logging names."""
    import win32com.client

    connection = win32com.client.Dispatch("ADODB.Connection")
    connection.Open(f"Provider=Microsoft.ACE.OLEDB.16.0;Mode=Read;Data Source={path}")
    references = [
        ("Workbooks", "WorkBookName", "WorkbookID"),
        ("Permissions", "WorkBookName", "WorkbookID"),
        ("Documents", "Workbook", "WorkbookID"),
        ("BatchList", "Workbookname", "WorkbookID"),
        ("Descriptions", "WorkbookName", "WorkbookID"),
        ("DescriptionDimensions", "WorkbookName", "WorkbookID"),
        ("DataLinks", "TargetWB", "TargetWorkbookID"),
        ("DataLinks", "SourceWB", "SourceWorkbookID"),
    ]
    try:
        available = {name.lower() for name in schema}
        for table, name, identity in references:
            if table.lower() not in available:
                continue
            records = connection.Execute(
                f"SELECT COUNT(*) FROM [{table}] AS r LEFT JOIN WorkbookIdentities AS i "
                f"ON r.[{identity}] = i.WorkbookID "
                f"WHERE (r.[{name}] & '') <> '' AND "
                f"(i.WorkbookID IS NULL OR (r.[{name}] & '') <> i.WorkbookName)"
            )[0]
            valid = records.Fields.Item(0).Value == 0
            records.Close()
            if not valid:
                return False
        return True
    finally:
        connection.Close()


def main():
    import win32com.client

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--database-copy", type=Path, action="append", default=[])
    parser.add_argument("--legacy-schema-source", type=Path)
    parser.add_argument("--output-parent", type=Path)
    parser.add_argument("--database-only", action="store_true")
    args = parser.parse_args()
    repo = Path(__file__).resolve().parents[1]
    output = Path(
        tempfile.mkdtemp(prefix="nb-access-identity-", dir=args.output_parent)
    )
    source = (repo / "vba/modules/DbCreateTables.bas").read_text("utf-8-sig")
    helpers = procedures(
        source,
        [
            "StripBrackets",
            "SqlBracket",
            "SqlString",
            "DbObjectName",
            "PrepareTable",
            "AddStrCol",
            "AddDateCol",
            "AttachTable",
            "Prepareindex",
            "IndexCol",
            "AttachIndex",
            "IndexExists",
            "CloseRecordsetSafely",
            "DBTableExists",
            "DBColumnExists",
            "DBColumnSize",
            "DropIndexIfExists",
            "RecreateUniqueIndex",
            "CreateTableWorkbookIdentities",
            "CreateTableSchemaMigrations",
            "AlterFieldLen",
        ],
    )
    helpers = source.split("Private Function StripBrackets", 1)[0] + helpers
    migration = (repo / "vba/modules/WorkbookIdentitySchema.bas").read_text("utf-8-sig")
    functions = procedures(
        (repo / "vba/modules/DbFunc.bas").read_text("utf-8-sig"),
        ["InQ", "InB", "NowFunction"],
    )
    execute = procedures(
        (repo / "vba/modules/DBInterface.bas").read_text("utf-8-sig"), ["DbExecute"]
    )
    # Inject one provider failure to verify diagnostics/retry, not to claim that
    # the reported Marshall Islands failure has been reproduced.
    execute = execute.replace(
        "    CurrentDB.DBCnn.Execute ssql, Recaff",
        """    If InjectProviderFailure And InStr(ssql, "[Permissions] ALTER COLUMN") > 0 Then
        InjectProviderFailure = False
        Err.Raise -2147467259, "Synthetic provider failure", "Record is deleted."
    End If
    CurrentDB.DBCnn.Execute ssql, Recaff""",
    )
    excel = win32com.client.DispatchEx("Excel.Application")
    book = None
    report = {"scope": "synthetic Access with production VBA migration", "cases": []}
    try:
        excel.EnableEvents = False
        excel.DisplayAlerts = False
        book = excel.Workbooks.Add()
        project = book.VBProject
        project.References.AddFromGuid("{2A75196C-D9EB-4129-B803-931327F72D5C}", 2, 8)
        project.References.AddFromGuid("{00000600-0000-0010-8000-00AA006D2EA4}", 2, 8)
        cls = project.VBComponents.Add(2)
        cls.Name = "ProbeDatabase"
        cls.CodeModule.AddFromString("""Option Explicit
Public DBCnn As ADODB.Connection
Public DBCat As ADOX.Catalog
Public DBType As Long
Public DbIsExch As Boolean
""")
        support = (
            """Option Explicit
Public CurrentDB As ProbeDatabase
Public Const Sqlexpress As Long = 1
Public Const mdb As Long = 2
Public Const accdb As Long = 3
Public ProbeLog As String
Public InjectProviderFailure As Boolean
Public Sub AddToSQLLog(ByVal Message As String)
ProbeLog = ProbeLog & Message & vbCrLf
End Sub
Public Sub RecordProbeMessage(ByVal Message As String, ByVal Style As Long, ByVal Title As String)
AddToSQLLog Message
End Sub
Public Sub OpenDb()
End Sub
Public Sub CloseDB()
End Sub
Public Function RunProbe(ByVal FilePath As String, ByVal InvalidReference As Boolean) As String
Dim FirstResult As Boolean
Dim Again As Boolean
Dim Snapshot As ADODB.Recordset
On Error GoTo Failed
Set CurrentDB = New ProbeDatabase
CurrentDB.DBType = accdb
Set CurrentDB.DBCat = New ADOX.Catalog
CurrentDB.DBCat.Create "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & FilePath
Set CurrentDB.DBCnn = CurrentDB.DBCat.ActiveConnection
CurrentDB.DBCnn.Execute "CREATE TABLE Workbooks (WorkBookName TEXT(50) NOT NULL, FilePath TEXT(255))"
CurrentDB.DBCnn.Execute "CREATE UNIQUE INDEX PrimaryIndex ON Workbooks (WorkBookName)"
CurrentDB.DBCnn.Execute "INSERT INTO Workbooks (WorkBookName, FilePath) VALUES ('Synthetic workbook', '!')"
CurrentDB.DBCnn.Execute "CREATE TABLE Permissions (WorkBookName TEXT(50), UserName TEXT(50))"
CurrentDB.DBCnn.Execute "INSERT INTO Permissions VALUES ('Synthetic workbook', 'test-user')"
CurrentDB.DBCnn.Execute "CREATE TABLE Documents (Workbook TEXT(50), DocumentName TEXT(50))"
CurrentDB.DBCnn.Execute "INSERT INTO Documents VALUES ('Synthetic workbook', 'test-document')"
CurrentDB.DBCnn.Execute "CREATE TABLE BatchList (Workbookname TEXT(50))"
CurrentDB.DBCnn.Execute "INSERT INTO BatchList VALUES ('Synthetic workbook')"
CurrentDB.DBCnn.Execute "CREATE TABLE Descriptions (WorkbookName TEXT(50), DataAreaName TEXT(50))"
CurrentDB.DBCnn.Execute "INSERT INTO Descriptions VALUES ('Synthetic workbook', 'area')"
CurrentDB.DBCnn.Execute "CREATE UNIQUE INDEX PrimaryIndex ON Descriptions (WorkbookName, DataAreaName)"
CurrentDB.DBCnn.Execute "CREATE TABLE DescriptionDimensions (WorkbookName TEXT(50), DataAreaName TEXT(50), DimensionNumber LONG)"
CurrentDB.DBCnn.Execute "INSERT INTO DescriptionDimensions VALUES ('Synthetic workbook', 'area', 1)"
CurrentDB.DBCnn.Execute "CREATE UNIQUE INDEX PrimaryIndex ON DescriptionDimensions (WorkbookName, DataAreaName, DimensionNumber)"
CurrentDB.DBCnn.Execute "CREATE TABLE DataLinks (KeyFamily TEXT(50), TargetWB TEXT(50), TargetDataArea TEXT(50), SourceWB TEXT(50), SourceDataArea TEXT(50))"
CurrentDB.DBCnn.Execute "INSERT INTO DataLinks VALUES ('test-key', 'Synthetic workbook', 'target-area', 'Source workbook', 'source-area')"
CurrentDB.DBCnn.Execute "CREATE UNIQUE INDEX PrimaryIndex ON DataLinks (KeyFamily, TargetWB, TargetDataArea, SourceWB, SourceDataArea)"
CurrentDB.DBCnn.Execute "INSERT INTO Workbooks (WorkBookName, FilePath) VALUES ('L''économie Å ', '!')"
CurrentDB.DBCnn.Execute "INSERT INTO DescriptionDimensions VALUES ('L''économie Å ', 'edge-area', 1)"
CurrentDB.DBCnn.Execute "INSERT INTO Documents VALUES (NULL, 'unlinked-null')"
CurrentDB.DBCnn.Execute "INSERT INTO Documents VALUES ('', 'unlinked-empty')"
InjectProviderFailure = InvalidReference
ProbeLog = ""
FirstResult = EnsureWorkbookIdentitySchema()
If InvalidReference Then
    If FirstResult Then Err.Raise vbObjectError + 7000, , "Invalid reference was accepted"
    If InStr(ProbeLog, "Step: Require identity column: Permissions.WorkbookID") = 0 Then Err.Raise vbObjectError + 7001, , "Missing precise failure step"
    Set Snapshot = CurrentDB.DBCnn.Execute("SELECT COUNT(*) FROM NADABASSchemaMigrations")
    If Snapshot.Fields(0).Value <> 0 Then Err.Raise vbObjectError + 7002, , "Failed migration marked complete"
    Snapshot.Close
    FirstResult = EnsureWorkbookIdentitySchema()
End If
If Not FirstResult Then Err.Raise vbObjectError + 7003, , "Migration failed"
Set Snapshot = CurrentDB.DBCnn.Execute("SELECT COUNT(*) FROM WorkbookIdentities")
If Snapshot.Fields(0).Value <> 3 Then Err.Raise vbObjectError + 7004, , "Identity count changed"
Snapshot.Close
Set Snapshot = CurrentDB.DBCnn.Execute("SELECT COUNT(*) FROM DescriptionDimensions WHERE WorkbookID = " & CStr(GetWorkbookID("L'économie Å ")))
If Snapshot.Fields(0).Value <> 1 Then Err.Raise vbObjectError + 7008, , "Accented or quoted workbook reference lost"
Snapshot.Close
Set Snapshot = CurrentDB.DBCnn.Execute("SELECT COUNT(*) FROM Documents WHERE WorkbookID IS NULL")
If Snapshot.Fields(0).Value <> 2 Then Err.Raise vbObjectError + 7009, , "Unlinked document changed"
Snapshot.Close
Again = EnsureWorkbookIdentitySchema()
If Not Again Then Err.Raise vbObjectError + 7005, , "Repeated migration failed"
Set Snapshot = CurrentDB.DBCnn.Execute("SELECT COUNT(*) FROM Workbooks WHERE WorkBookName = 'Synthetic workbook' AND FilePath = '!'")
If Snapshot.Fields(0).Value <> 1 Then Err.Raise vbObjectError + 7006, , "Workbook data changed"
Snapshot.Close
Set Snapshot = CurrentDB.DBCnn.Execute("SELECT COUNT(*) FROM NADABASSchemaMigrations")
If Snapshot.Fields(0).Value <> 1 Then Err.Raise vbObjectError + 7007, , "Migration marker duplicated"
Snapshot.Close
RunProbe = "True" & vbCrLf & ProbeLog
CurrentDB.DBCnn.Close
Set CurrentDB = Nothing
Exit Function
Failed:
RunProbe = "ERROR " & Err.Number & ": " & Err.Description & vbCrLf & ProbeLog
On Error Resume Next
CurrentDB.DBCnn.Close
Set CurrentDB = Nothing
End Function
"""
            + functions
            + "\n"
            + execute
            + """
Public Function RunExistingCopy(ByVal FilePath As String) As String
Dim Success As Boolean
Dim Counts As Object
Dim TableName As Variant
Dim Snapshot As ADODB.Recordset
On Error GoTo Failed
InjectProviderFailure = False
ProbeLog = ""
Set CurrentDB = New ProbeDatabase
CurrentDB.DBType = accdb
Set CurrentDB.DBCnn = New ADODB.Connection
CurrentDB.DBCnn.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & FilePath
Set CurrentDB.DBCat = New ADOX.Catalog
Set CurrentDB.DBCat.ActiveConnection = CurrentDB.DBCnn
Set Counts = CreateObject("Scripting.Dictionary")
For Each TableName In Array("Workbooks", "Permissions", "Documents", "BatchList", "Descriptions", "DescriptionDimensions", "DataLinks")
    If DBTableExists(CStr(TableName)) Then
        Set Snapshot = CurrentDB.DBCnn.Execute("SELECT * FROM [" & TableName & "]")
        Counts.Add CStr(TableName), EnumeratedRows(Snapshot)
        Snapshot.Close
    End If
Next TableName
Success = EnsureWorkbookIdentitySchema()
For Each TableName In Counts.Keys
    Set Snapshot = CurrentDB.DBCnn.Execute("SELECT * FROM [" & TableName & "]")
    If EnumeratedRows(Snapshot) <> Counts(TableName) Then Err.Raise vbObjectError + 7020, , "Row count changed in " & TableName
    Snapshot.Close
Next TableName
RunExistingCopy = CStr(Success) & vbCrLf & ProbeLog
CurrentDB.DBCnn.Close
Set CurrentDB = Nothing
Exit Function
Failed:
RunExistingCopy = "ERROR " & Err.Number & ": " & Err.Description & vbCrLf & ProbeLog
On Error Resume Next
CurrentDB.DBCnn.Close
Set CurrentDB = Nothing
End Function
Private Function EnumeratedRows(ByVal Records As ADODB.Recordset) As Long
Do While Not Records.EOF
    EnumeratedRows = EnumeratedRows + 1
    Records.MoveNext
Loop
End Function
"""
        )
        modules = [
            ("ProbeSupport", support),
            ("DbCreateTables", helpers),
            ("WorkbookIdentitySchema", migration),
        ]
        if args.legacy_schema_source:
            old = args.legacy_schema_source.read_text("utf-8-sig")
            legacy = old.split("Private Function StripBrackets", 1)[0] + procedures(
                old,
                [
                    "StripBrackets",
                    "SqlBracket",
                    "SqlString",
                    "DbObjectName",
                    "PrepareTable",
                    "AddStrCol",
                    "AddLongCol",
                    "AddDateCol",
                    "AddMemoCol",
                    "AttachTable",
                    "Prepareindex",
                    "IndexCol",
                    "AttachIndex",
                    "IndexExists",
                    "DropTable",
                    "CreateWorkbookTable",
                    "CreateTablePermissions",
                    "CreateTableDocuments",
                    "CreateTableBatchList",
                    "CreateTableDescriptions",
                    "CreateTableDescriptionsDimensions",
                    "CreateTableDataLinks",
                ],
            )
            legacy = legacy.replace(
                'Attribute VB_Name = "DbCreateTables"',
                'Attribute VB_Name = "LegacyFixture"',
            )
            modules.append(("LegacyFixture", legacy))
            modules.append(
                (
                    "LegacyRunner",
                    """Option Explicit
Public Function CreateLegacyFixture(ByVal FilePath As String) As String
On Error GoTo Failed
Set CurrentDB = New ProbeDatabase
CurrentDB.DBType = accdb
Set CurrentDB.DBCat = New ADOX.Catalog
CurrentDB.DBCat.Create "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & FilePath
Set CurrentDB.DBCnn = CurrentDB.DBCat.ActiveConnection
LegacyFixture.CreateWorkbookTable
LegacyFixture.CreateTablePermissions
LegacyFixture.CreateTableDocuments
LegacyFixture.CreateTableBatchList
LegacyFixture.CreateTableDescriptions
LegacyFixture.CreateTableDescriptionsDimensions
LegacyFixture.CreateTableDataLinks
CurrentDB.DBCnn.Execute "INSERT INTO Workbooks (WorkBookName, Path, Title) VALUES ('Legacy workbook', '!', 'Legacy title')"
CurrentDB.DBCnn.Execute "INSERT INTO Permissions (WorkBookName, [User]) VALUES ('Legacy workbook', 'test-user')"
CurrentDB.DBCnn.Execute "INSERT INTO Documents ([Name], [Path], Workbook) VALUES ('test-document', '!', 'Legacy workbook')"
CurrentDB.DBCnn.Execute "INSERT INTO BatchList (Listname, ItemNo, Workbookname) VALUES ('test-batch', 1, 'Legacy workbook')"
CurrentDB.DBCnn.Execute "INSERT INTO Descriptions (WorkbookName, DataAreaName, TableName, GetPut) VALUES ('Legacy workbook', 'area', 'test-table', 'Get')"
CurrentDB.DBCnn.Execute "INSERT INTO DescriptionDimensions (WorkbookName, DataAreaName, DimensionNumber, DimensionName) VALUES ('Legacy workbook', 'area', 1, 'dimension')"
CurrentDB.DBCnn.Execute "INSERT INTO DataLinks (KeyFamily, TargetWB, TargetDataArea, SourceWB, SourceDataArea) VALUES ('test-key', 'Legacy workbook', 'target', 'Source workbook', 'source')"
CurrentDB.DBCnn.Close
Set CurrentDB = Nothing
CreateLegacyFixture = "True"
Exit Function
Failed:
CreateLegacyFixture = "ERROR " & Err.Number & ": " & Err.Description
On Error Resume Next
CurrentDB.DBCnn.Close
Set CurrentDB = Nothing
End Function
""",
                )
            )
        for name, code in modules:
            module = project.VBComponents.Add(1)
            module.Name = name
            code = re.sub(r"^Attribute .*\n?", "", code, flags=re.M)
            # Capture UI messages only; production migration and SQL are unchanged.
            module.CodeModule.AddFromString(
                code.replace("MsgBox ", "RecordProbeMessage ")
            )
        synthetic_cases = [
            ("legacy-all-tables", False),
            ("partial-failure-retry", True),
            ("switch-to-another-database", False),
        ]
        for name, invalid in ([] if args.database_only else synthetic_cases):
            result = excel.Run(
                f"'{book.Name}'!ProbeSupport.RunProbe",
                str(output / f"{name}.accdb"),
                invalid,
            )
            report["cases"].append(
                {
                    "name": name,
                    "passed": str(result).startswith("True"),
                    "result": result,
                }
            )
        if args.legacy_schema_source:
            legacy_db = output / "from-legacy-schema.accdb"
            created = str(
                excel.Run(
                    f"'{book.Name}'!LegacyRunner.CreateLegacyFixture", str(legacy_db)
                )
            )
            if created != "True":
                raise RuntimeError(created)
            for attempt in range(2):
                result = str(
                    excel.Run(
                        f"'{book.Name}'!ProbeSupport.RunExistingCopy", str(legacy_db)
                    )
                )
                report["cases"].append(
                    {
                        "name": f"legacy-schema-migration-{attempt}",
                        "passed": result.startswith("True"),
                        "result": result,
                    }
                )
        for number, original in enumerate(args.database_copy):
            original = original.resolve(strict=True)
            if (
                original.suffix.lower() != ".accdb"
                or original.with_suffix(".laccdb").exists()
            ):
                raise ValueError(
                    "Use an unlocked, explicitly identified Access test database"
                )
            before = hashlib.sha256(original.read_bytes()).hexdigest()
            copied = output / f"existing-copy-{number}.accdb"
            shutil.copyfile(original, copied)
            if hashlib.sha256(copied.read_bytes()).hexdigest() != before:
                raise ValueError("Test database changed during copying")
            schema, original_rows = legacy_snapshot(copied)
            result = str(
                excel.Run(f"'{book.Name}'!ProbeSupport.RunExistingCopy", str(copied))
            )
            _, migrated_rows = legacy_snapshot(copied, schema)
            preserved = original_rows == migrated_rows
            identities_valid = verify_identity_references(copied, schema)
            repeated = str(
                excel.Run(f"'{book.Name}'!ProbeSupport.RunExistingCopy", str(copied))
            )
            _, repeated_rows = legacy_snapshot(copied, schema)
            passed = (
                result.startswith("True")
                and repeated.startswith("True")
                and preserved
                and identities_valid
                and repeated_rows == original_rows
            )
            if hashlib.sha256(original.read_bytes()).hexdigest() != before:
                raise ValueError("Original test database changed during the test")
            # Do not report workbook names or SQL containing user data.
            lines = result.splitlines()
            report["cases"].append(
                {
                    "name": f"existing-copy-{number}",
                    "originalSha256": before,
                    "passed": passed,
                    "legacyDataPreserved": preserved,
                    "identityReferencesValid": identities_valid,
                    "tablesChecked": len(schema),
                    "repeatPassed": repeated.startswith("True")
                    and repeated_rows == original_rows,
                    "result": [
                        line
                        for line in lines
                        if line.startswith(("Workbook identity migration", "Step:"))
                    ],
                    "errors": re.findall(r"(?:ERROR|Error) (-?\d+)", result),
                    "status": "passed" if passed else "failed",
                }
            )
        report["passed"] = all(case["passed"] for case in report["cases"])
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        excel.Quit()
        (output / "report.json").write_text(
            json.dumps(report, indent=2), encoding="utf-8"
        )
        print(output / "report.json")
    print(json.dumps(report, indent=2))
    return 0 if report.get("passed") else 1


if __name__ == "__main__":
    raise SystemExit(main())
