Attribute VB_Name = "WorkbookIdentitySchema"
Option Explicit
Option Private Module

Public Const WORKBOOK_NAME_MAX_LENGTH As Long = 255

Private Const WORKBOOK_ID_MIGRATION As String = "WorkbookIdentityV1"

Public Function EnsureWorkbookIdentitySchema() As Boolean

    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo ErrorHandler

    EnsureWorkbookIdentitySchema = False

    If CurrentDB.DbIsExch Then
        EnsureWorkbookIdentitySchema = True
        Exit Function
    End If

    If Not DBTableExists("Workbooks") Then
        EnsureWorkbookIdentitySchema = True
        Exit Function
    End If

    OpenDb

    CreateTableSchemaMigrations

    If WorkbookIdentityMigrationApplied Then
        EnsureWorkbookIdentitySchema = True
        GoTo CleanExit
    End If

    CreateTableWorkbookIdentities
    EnsureWorkbookIDColumns
    BackfillWorkbookIdentityData
    RebuildWorkbookIdentityIndexes
    RequireWorkbookIDColumns

    WidenLegacyWorkbookNameColumns
    CreateWorkbookIdentityIndexes

    DbExecute "INSERT INTO [NADABASSchemaMigrations] " & _
              "([MigrationName], [AppliedAt]) VALUES (" & _
              InQ(WORKBOOK_ID_MIGRATION) & ", " & NowFunction & ")"

    EnsureWorkbookIdentitySchema = True

CleanExit:
    CloseDB
    Exit Function

ErrorHandler:
    errorNumber = err.Number
    errorDescription = err.Description

    AddToSQLLog "Workbook identity migration failed" & vbCrLf & _
                errorNumber & ":" & errorDescription

    MsgBox "NADABAS could not upgrade workbook references in the database." & _
           vbCrLf & vbCrLf & _
           "No workbook data was deleted. Ask a database administrator to " & _
           "open the database with schema-change permission and try again." & _
           vbCrLf & vbCrLf & _
           "Error " & errorNumber & ": " & errorDescription, _
           vbCritical, "NADABAS"

    Resume CleanExit

End Function

Public Function GetWorkbookID(WorkbookName As String) As Long

    Dim rs As ADODB.Recordset
    Dim sql As String

    On Error GoTo CleanExit

    If Len(WorkbookName) = 0 Then Exit Function
    If Not DBTableExists("WorkbookIdentities") Then Exit Function

    sql = "SELECT [WorkbookID] FROM [WorkbookIdentities] " & _
          "WHERE [WorkbookName] = " & InQ(WorkbookName)

    Set rs = CurrentDB.DBCnn.Execute(sql)

    If Not rs.EOF Then
        GetWorkbookID = CLng(rs.fields("WorkbookID").value)
    End If

CleanExit:
    CloseRecordsetSafely rs

End Function

Public Function GetOrCreateWorkbookID(WorkbookName As String) As Long

    Dim rs As ADODB.Recordset
    Dim errorNumber As Long
    Dim errorDescription As String

    If Len(WorkbookName) = 0 Then Exit Function
    If Len(WorkbookName) > WORKBOOK_NAME_MAX_LENGTH Then
        err.Raise vbObjectError + 6102, "GetOrCreateWorkbookID", _
                  "Workbook names cannot exceed " & _
                  CStr(WORKBOOK_NAME_MAX_LENGTH) & " characters."
    End If

    If Not DBTableExists("WorkbookIdentities") Then
        CreateTableWorkbookIdentities
    End If

    GetOrCreateWorkbookID = GetWorkbookID(WorkbookName)
    If GetOrCreateWorkbookID <> 0 Then Exit Function

    On Error GoTo InsertFailed

    DbExecute "INSERT INTO [WorkbookIdentities] ([WorkbookName]) VALUES (" & _
              InQ(WorkbookName) & ")"

    Set rs = CurrentDB.DBCnn.Execute("SELECT @@IDENTITY AS [WorkbookID]")
    If Not rs.EOF Then
        GetOrCreateWorkbookID = CLng(rs.fields("WorkbookID").value)
    End If

CleanExit:
    CloseRecordsetSafely rs
    Exit Function

InsertFailed:
    errorNumber = err.Number
    errorDescription = err.Description

    GetOrCreateWorkbookID = GetWorkbookID(WorkbookName)
    If GetOrCreateWorkbookID = 0 Then
        err.Raise errorNumber, "GetOrCreateWorkbookID", errorDescription
    End If

    Resume CleanExit

End Function

Public Sub RenameWorkbookIdentity(WorkbookID As Long, NewName As String)

    Dim ExistingID As Long

    If WorkbookID = 0 Then
        err.Raise vbObjectError + 6100, "RenameWorkbookIdentity", _
                  "The workbook has no database identity."
    End If

    If Len(NewName) = 0 Or Len(NewName) > WORKBOOK_NAME_MAX_LENGTH Then
        err.Raise vbObjectError + 6102, "RenameWorkbookIdentity", _
                  "Workbook names must contain between 1 and " & _
                  CStr(WORKBOOK_NAME_MAX_LENGTH) & " characters."
    End If

    ExistingID = GetWorkbookID(NewName)

    If ExistingID <> 0 And ExistingID <> WorkbookID Then
        err.Raise vbObjectError + 6101, "RenameWorkbookIdentity", _
                  "The workbook name is assigned to another workbook identity: " & _
                  NewName
    End If

    DbExecute "UPDATE [WorkbookIdentities] SET [WorkbookName] = " & _
              InQ(NewName) & " WHERE [WorkbookID] = " & CStr(WorkbookID)

End Sub

Private Function WorkbookIdentityMigrationApplied() As Boolean

    Dim rs As ADODB.Recordset

    On Error GoTo CleanExit

    Set rs = CurrentDB.DBCnn.Execute( _
        "SELECT [MigrationName] FROM [NADABASSchemaMigrations] " & _
        "WHERE [MigrationName] = " & InQ(WORKBOOK_ID_MIGRATION))

    WorkbookIdentityMigrationApplied = Not rs.EOF

CleanExit:
    CloseRecordsetSafely rs

End Function

Private Sub CloseRecordsetSafely(ByRef rs As ADODB.Recordset)

    On Error Resume Next

    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing

End Sub

Private Sub EnsureWorkbookIDColumns()

    EnsureLongColumn "Workbooks", "WorkbookID"
    EnsureLongColumn "Permissions", "WorkbookID"
    EnsureLongColumn "Documents", "WorkbookID"
    EnsureLongColumn "BatchList", "WorkbookID"
    EnsureLongColumn "Descriptions", "WorkbookID"
    EnsureLongColumn "DescriptionDimensions", "WorkbookID"
    EnsureLongColumn "DataLinks", "TargetWorkbookID"
    EnsureLongColumn "DataLinks", "SourceWorkbookID"

End Sub

Private Sub EnsureLongColumn(TableName As String, ColumnName As String)

    If Not DBTableExists(TableName) Then Exit Sub
    If DBColumnExists(TableName, ColumnName) Then Exit Sub

    Select Case CurrentDB.DBType
        Case Sqlexpress
            DbExecute "ALTER TABLE " & InB(TableName) & " ADD " & _
                      InB(ColumnName) & " INT NULL"
        Case accdb, mdb
            DbExecute "ALTER TABLE " & InB(TableName) & " ADD COLUMN " & _
                      InB(ColumnName) & " LONG"
        Case Else
            err.Raise vbObjectError + 6103, "EnsureLongColumn", _
                      "Unsupported database type: " & CStr(CurrentDB.DBType)
    End Select

End Sub

Private Sub RequireWorkbookIDColumns()

    RequireLongColumn "Workbooks", "WorkbookID"
    RequireLongColumn "Permissions", "WorkbookID"
    RequireLongColumn "BatchList", "WorkbookID"
    RequireLongColumn "Descriptions", "WorkbookID"
    RequireLongColumn "DescriptionDimensions", "WorkbookID"
    RequireLongColumn "DataLinks", "TargetWorkbookID"
    RequireLongColumn "DataLinks", "SourceWorkbookID"

End Sub

Private Sub RequireLongColumn(TableName As String, ColumnName As String)

    If Not DBTableExists(TableName) Then Exit Sub

    Select Case CurrentDB.DBType
        Case Sqlexpress
            DbExecute "ALTER TABLE " & InB(TableName) & " ALTER COLUMN " & _
                      InB(ColumnName) & " INT NOT NULL"
        Case accdb, mdb
            DbExecute "ALTER TABLE " & InB(TableName) & " ALTER COLUMN " & _
                      InB(ColumnName) & " LONG NOT NULL"
        Case Else
            err.Raise vbObjectError + 6104, "RequireLongColumn", _
                      "Unsupported database type: " & CStr(CurrentDB.DBType)
    End Select

End Sub

Private Sub BackfillWorkbookIdentityData()

    BackfillWorkbookIDs "Workbooks", "WorkBookName", "WorkbookID"
    BackfillWorkbookIDs "Permissions", "WorkBookName", "WorkbookID"
    BackfillWorkbookIDs "Documents", "Workbook", "WorkbookID"
    BackfillWorkbookIDs "BatchList", "Workbookname", "WorkbookID"
    BackfillWorkbookIDs "Descriptions", "WorkbookName", "WorkbookID"
    BackfillWorkbookIDs "DescriptionDimensions", "WorkbookName", "WorkbookID"
    BackfillWorkbookIDs "DataLinks", "TargetWB", "TargetWorkbookID"
    BackfillWorkbookIDs "DataLinks", "SourceWB", "SourceWorkbookID"

End Sub

Private Sub BackfillWorkbookIDs(TableName As String, NameColumn As String, _
                                IDColumn As String)

    Dim rs As ADODB.Recordset
    Dim WorkbookNames As Collection
    Dim NameValue As Variant
    Dim WorkbookName As String
    Dim WorkbookID As Long
    Dim sql As String

    If Not DBTableExists(TableName) Then Exit Sub

    sql = "SELECT DISTINCT " & InB(NameColumn) & " FROM " & InB(TableName) & _
          " WHERE " & InB(IDColumn) & " IS NULL"

    Set WorkbookNames = New Collection
    Set rs = CurrentDB.DBCnn.Execute(sql)

    Do While Not rs.EOF
        If Not IsNull(rs.fields(0).value) Then
            WorkbookName = CStr(rs.fields(0).value)
            If Len(WorkbookName) <> 0 Then
                WorkbookNames.Add WorkbookName
            End If
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    For Each NameValue In WorkbookNames
        WorkbookName = CStr(NameValue)
        WorkbookID = GetOrCreateWorkbookID(WorkbookName)
        DbExecute "UPDATE " & InB(TableName) & " SET " & _
                  InB(IDColumn) & " = " & CStr(WorkbookID) & _
                  " WHERE " & InB(IDColumn) & " IS NULL AND " & _
                  InB(NameColumn) & " = " & InQ(WorkbookName)
    Next NameValue

End Sub

Private Sub RebuildWorkbookIdentityIndexes()

    DropIndexIfExists "Workbooks", "PrimaryIndex"

    If DBTableExists("Descriptions") Then
        DropIndexIfExists "Descriptions", "PrimaryIndex"
    End If

    If DBTableExists("DescriptionDimensions") Then
        DropIndexIfExists "DescriptionDimensions", "PrimaryIndex"
    End If

    If DBTableExists("DataLinks") Then
        DropIndexIfExists "DataLinks", "PrimaryIndex"
    End If

End Sub

Private Sub WidenLegacyWorkbookNameColumns()

    EnsureWorkbookNameLength "Workbooks", "WorkBookName"
    EnsureWorkbookNameLength "Permissions", "WorkBookName"
    EnsureWorkbookNameLength "Documents", "Workbook"
    EnsureWorkbookNameLength "BatchList", "Workbookname"
    EnsureWorkbookNameLength "Descriptions", "WorkbookName"
    EnsureWorkbookNameLength "DescriptionDimensions", "WorkbookName"
    EnsureWorkbookNameLength "DataLinks", "TargetWB"
    EnsureWorkbookNameLength "DataLinks", "SourceWB"

End Sub

Private Sub EnsureWorkbookNameLength(TableName As String, ColumnName As String)

    If Not DBTableExists(TableName) Then Exit Sub
    If Not DBColumnExists(TableName, ColumnName) Then Exit Sub

    If DBColumnSize(TableName, ColumnName) < WORKBOOK_NAME_MAX_LENGTH Then
        AlterFieldLen TableName, ColumnName, WORKBOOK_NAME_MAX_LENGTH
    End If

End Sub

Private Sub CreateWorkbookIdentityIndexes()

    RecreateUniqueIndex "PrimaryIndex", "WorkbookIdentities", "WorkbookID"
    RecreateUniqueIndex "WorkbookNameIndex", "WorkbookIdentities", "WorkbookName"

    RecreateUniqueIndex "PrimaryIndex", "Workbooks", "WorkbookID"
    RecreateUniqueIndex "WorkbookNameIndex", "Workbooks", "WorkBookName"

    If DBTableExists("Descriptions") Then
        RecreateUniqueIndex "PrimaryIndex", "Descriptions", _
                            "WorkbookID", "DataAreaName"
    End If

    If DBTableExists("DescriptionDimensions") Then
        RecreateUniqueIndex "PrimaryIndex", "DescriptionDimensions", _
                            "WorkbookID", "DataAreaName", "DimensionNumber"
    End If

    If DBTableExists("DataLinks") Then
        RecreateUniqueIndex "PrimaryIndex", "DataLinks", _
                            "KeyFamily", "TargetWorkbookID", _
                            "TargetDataArea", "SourceWorkbookID", _
                            "SourceDataArea"
    End If

End Sub
