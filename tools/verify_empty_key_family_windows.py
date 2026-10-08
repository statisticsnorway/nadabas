"""Run the exact production row guard in scratch Excel against disposable ACE tables."""

import json
import tempfile
from pathlib import Path

import win32com.client as com

root = Path(__file__).resolve().parents[1]
(root / ".codex-build").mkdir(exist_ok=True)
fixture = Path(tempfile.mkdtemp(prefix="empty-guard-", dir=root / ".codex-build"))
db = fixture / "fixture.accdb"
catalog = com.Dispatch("ADOX.Catalog")
catalog.Create(f"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={db};")
conn = catalog.ActiveConnection
conn.Execute("CREATE TABLE [EmptyFamily] ([K] TEXT(20), [Value] DOUBLE)")
conn.Execute("CREATE TABLE [PopulatedFamily] ([K] TEXT(20), [Value] DOUBLE)")
conn.Execute("INSERT INTO [PopulatedFamily] ([K]) VALUES ('null-value-row')")
excel = com.DispatchEx("Excel.Application")
book = None
results = {}
try:
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    book = excel.Workbooks.Add()
    holder = book.VBProject.VBComponents.Add(2)
    holder.Name = "ConnectionHolder"
    holder.CodeModule.AddFromString(
        "Option Explicit\nPublic DBCnn As Object\nPublic DBCat As Object\n"
    )
    module = book.VBProject.VBComponents.Add(1)
    module.Name = "EmptyGuardSmoke"
    policy = (root / "vba/modules/KeyFamilyEditPolicy.bas").read_text(encoding="utf-8")
    helper = policy[policy.index("Public Function TableIsEmptyForStructure(") :]
    module.CodeModule.AddFromString(
        "Option Explicit\nPublic Held As Object, HeldCat As Object\n" + helper + """
Public Function CheckTable(ByVal DbPath As String, ByVal TableName As String, ByVal Closed As Boolean) As Boolean
Dim cnn As Object, reason As String
Set cnn = CreateObject("ADODB.Connection")
If Not Closed Then cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DbPath & ";"
CheckTable = TableIsEmptyForStructure(cnn, TableName, reason)
If cnn.State <> 0 Then cnn.Close
End Function
Public Function CheckExclusive(ByVal DbPath As String) As Boolean
Dim cnn As Object, cat As Object, originalMode As Long
Set cnn = CreateObject("ADODB.Connection")
Set cat = CreateObject("ADOX.Catalog")
cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DbPath & "; Jet OLEDB:Database;"
originalMode = cnn.Mode
Set cat.ActiveConnection = cnn
On Error GoTo Rejected
ReopenStructureConnection cnn, cat, "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DbPath & "; Jet OLEDB:Database;", 12
Set Held = cnn
Set HeldCat = cat
CheckExclusive = True
Exit Function
Rejected:
CheckExclusive = False
If cnn.State <> 0 Then cnn.Close
End Function
Public Sub ReleaseExclusive()
Set HeldCat = Nothing
If Not Held Is Nothing Then Held.Close
Set Held = Nothing
End Sub
Public Function CheckMemberConnection(ByVal DbPath As String) As String
Dim db As New ConnectionHolder, originalMode As Long, connectText As String
On Error GoTo Failed
connectText = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DbPath & "; Jet OLEDB:Database;"
Set db.DBCnn = CreateObject("ADODB.Connection")
Set db.DBCat = CreateObject("ADOX.Catalog")
db.DBCnn.Open connectText
Set db.DBCat.ActiveConnection = db.DBCnn
originalMode = db.DBCnn.Mode
ReopenStructureConnection db.DBCnn, db.DBCat, connectText, 12
db.DBCnn.BeginTrans
db.DBCnn.Execute "CREATE TABLE [MemberProbe] ([K] TEXT(20))"
db.DBCat.Tables.Refresh
db.DBCat.Tables("MemberProbe").Name = "MemberProbeRenamed"
db.DBCnn.RollbackTrans
ReopenStructureConnection db.DBCnn, db.DBCat, connectText, originalMode
db.DBCnn.BeginTrans
db.DBCnn.RollbackTrans
CheckMemberConnection = "PASS"
GoTo CleanUp
Failed:
CheckMemberConnection = CStr(Err.Number) & ": " & Err.Description
CleanUp:
On Error Resume Next
Set db.DBCat = Nothing
If Not db.DBCnn Is Nothing Then db.DBCnn.Close
End Function
"""
    )
    for table, closed, expected in [
        ("EmptyFamily", False, True),
        ("PopulatedFamily", False, False),
        ("MissingFamily", False, False),
        ("EmptyFamily", True, False),
    ]:
        actual = bool(excel.Run(f"'{book.Name}'!CheckTable", str(db), table, closed))
        assert actual == expected, (table, closed, actual)
        results[f"{table}/closed={closed}"] = actual
    # A row inserted after opening an editor must invalidate its prior empty result.
    conn.BeginTrans()
    conn.Execute("INSERT INTO [EmptyFamily] ([K]) VALUES ('late-row')")
    conn.CommitTrans()
    assert not excel.Run(f"'{book.Name}'!CheckTable", str(db), "EmptyFamily", False)
    results["late_insert_rejected"] = True
    # The same connection/catalogue used for DDL must undo create and rename.
    conn.BeginTrans()
    conn.Execute("CREATE TABLE [DraftFamily] ([NewK] TEXT(30), [Value] DOUBLE)")
    catalog.Tables.Refresh()
    catalog.Tables("EmptyFamily").Name = "RecoveryFamily"
    recordset = conn.Execute("SELECT TOP 1 1 FROM [RecoveryFamily]")[0]
    assert not recordset.EOF
    recordset.Close()
    conn.RollbackTrans()
    catalog.Tables.Refresh()
    names = [table.Name for table in catalog.Tables if table.Type == "TABLE"]
    assert (
        "EmptyFamily" in names
        and "RecoveryFamily" not in names
        and "DraftFamily" not in names
    )
    rs = conn.Execute("SELECT COUNT(*) FROM [EmptyFamily]")[0]
    assert rs.Fields(0).Value == 1
    rs.Close()
    results["rollback_preserves_late_row"] = True
    # A shared connection, including an async writer, must prevent the upgrade.
    conn.Execute("INSERT INTO [PopulatedFamily] ([K]) VALUES ('async-write')")
    exclusive_result = excel.Run(f"'{book.Name}'!CheckExclusive", str(db))
    assert not exclusive_result
    results["shared_connection_prevents_edit"] = True
    conn.Close()
    assert excel.Run(f"'{book.Name}'!CheckExclusive", str(db))
    outsider = com.Dispatch("ADODB.Connection")
    try:
        outsider.Open(f"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={db};")
    except Exception:
        pass
    else:
        outsider.Close()
        raise AssertionError("Exclusive Excel connection allowed another process")
    finally:
        excel.Run(f"'{book.Name}'!ReleaseExclusive")
    results["exclusive_connection_blocks_concurrent_open"] = True
    results["legacy_nadabas_connection_string_reopens_without_isam_error"] = True
    member_result = excel.Run(f"'{book.Name}'!CheckMemberConnection", str(db))
    print("Class-member connection check:", member_result, flush=True)
    assert member_result == "PASS"
    results["class_member_connection_transaction_and_restore"] = True
    (fixture / "results.json").write_text(
        json.dumps(results, indent=2), encoding="utf-8"
    )
    print(json.dumps(results, indent=2), flush=True)
finally:
    if book is not None:
        book.Close(SaveChanges=False)
    excel.Quit()
    if conn.State:
        conn.Close()
