Attribute VB_Name = "cmdConvertDB"
Option Explicit
Option Private Module




'
'   ********************************************************************************************
'   *                                                                                          *
'   *      The purpose of this module is to transer an NADABS acc database to SQL-server       *
'   *      Or Vica verca                                                                       *
'   *      Before running this module, the acc-database must exist                             *
'   *      and a empty SQL database should be prepared, i.e. having no tables at all           *
'   *      (or the other way round)                                                            *
'   *                                                                                          *
'   *      When NADABAS opens select the acc-database                                          *
'   *      This module then opens the sql-server DB                                            *
'   *                                                                                          *
'   *      If the SQL server DB is not empty, you may proceed.                                 *
'   *      In this case all tables are deleted from targer before continuing.
'   *                                                                                          *
'   ********************************************************************************************
'
Dim TargetDB As clsDB

Public Sub AccToSql()
' *******************
' called from Ribbon
' *******************

    Dim dbs As clsDB

    dlgAccToSQL.Show vbModal
    If dlgAccToSQL.cancel Then Exit Sub

    Set TargetDB = New clsDB

    '
    ' Temporarily point CurrentDB to TargetDB while user enters SQL connection data
    '
    Set CurrentDB = TargetDB
    dlgConnectString.Initialize (0)
    dlgConnectString.Show vbModal

    If dlgConnectString.cancel Then
        Set CurrentDB = BaseDb
        Exit Sub
    End If

    '
    ' Store source database
    '
    Set dbs = BaseDb

    '
    ' Build and test SQL connection
    '
    Set CurrentDB = TargetDB
    MakeconnectString

    If ConnectToSql = False Then
        Set BaseDb = dbs
        Set CurrentDB = BaseDb
        Exit Sub
    End If

    TargetDB.DBType = Sqlexpress

    '
    ' Return to Access source before transfer starts
    '
    Dim SourceBasePath As String

    Set BaseDb = dbs
    Set CurrentDB = BaseDb

    SourceBasePath = GetSourceBasePathForSqlCopy()

    If Not TransferData() Then

    Set BaseDb = dbs
    Set CurrentDB = BaseDb

        MsgBox "The database conversion was cancelled or failed." & vbCrLf & _
            "The SQL database has not been registered as a completed conversion.", _
            vbExclamation, "NADABAS conversion"

    Exit Sub

    End If

'
' SQL databases do not have a file-system location like Access databases.
' Therefore BasePath = "!" cannot be used in the SQL copy.
' Replace it with the actual source base folder.
'
    If Not SetTargetSqlBasePath(SourceBasePath) Then

        MsgBox "The database was copied to SQL Server, but NADABAS could not update the base folder." & vbCrLf & vbCrLf & _
           "You may need to set the base folder manually before opening workbooks.", _
           vbExclamation, "NADABAS conversion"

    End If

    If Not DatabaseAlreadyRegistered(TargetDB) Then
        Databases.Add TargetDB
    End If

    Set BaseDb = dbs
    Set CurrentDB = BaseDb

    SaveDBSettings

    Debug.Print "About to save SQL database:"
    Debug.Print "Registered SQL DB: " & TargetDB.DbDisplayName
    Debug.Print "Source: " & TargetDB.DBSource
    Debug.Print "Catalog: " & TargetDB.DBCatalog
    Debug.Print "DBType: " & TargetDB.DBType

    If MsgBox(GetMsg("M042") & vbCrLf & vbCrLf & _
            "The SQL database has been registered in NADABAS." & vbCrLf & _
            "NADABAS should now be restarted before using the SQL database." & vbCrLf & vbCrLf & _
            "Do you want to close Excel now?", _
            vbYesNo + vbInformation, "NADABAS conversion") = vbYes Then

        Application.quit

    End If
    ' "Database has been copied to SQL Server"

End Sub


Public Sub SqlToAcc()
' *******************
' called from Ribbon
' *******************

Dim dbs As clsDB
   dlgSQLToAcc.Show vbModal
   If dlgSQLToAcc.cancel Then Exit Sub

   Set TargetDB = New clsDB

   Set CurrentDB = TargetDB
   If GetAccDBName = False Then
      Set CurrentDB = BaseDb
      Exit Sub
   End If

   Set dbs = BaseDb
   Set BaseDb = TargetDB
   MakeAccConnectString
   If ConnectToSql = False Then
      Set BaseDb = dbs
      Set CurrentDB = BaseDb
      Exit Sub
   End If

   TargetDB.DBType = accdb
   Set BaseDb = dbs

   TransferData
   MsgBox GetMsg("M043") & TargetDB.DBFullName, vbOKOnly    'Database has been copied to  "
End Sub

Private Function DatabaseAlreadyRegistered(dbToFind As clsDB) As Boolean

    Dim dbx As clsDB

    DatabaseAlreadyRegistered = False

    For Each dbx In Databases

        If dbx.DBType = dbToFind.DBType Then

            Select Case dbToFind.DBType

                Case Sqlexpress

                    If LCase(Trim(dbx.DBSource)) = LCase(Trim(dbToFind.DBSource)) _
                       And LCase(Trim(dbx.DBCatalog)) = LCase(Trim(dbToFind.DBCatalog)) Then

                        DatabaseAlreadyRegistered = True
                        Exit Function

                    End If

                Case accdb, mdb

                    If LCase(Trim(dbx.DBFullName)) = LCase(Trim(dbToFind.DBFullName)) Then

                        DatabaseAlreadyRegistered = True
                        Exit Function

                    End If

            End Select

        End If

    Next dbx

End Function

Public Sub DOBackupSQL(FName As String)
'
' This is the backup function, that actually simply convert fra SQL to a new Accdb database fro backup.
'
 Dim dbs As clsDB
        Set TargetDB = New clsDB
        TargetDB.DBFullName = FName
        TargetDB.DBType = accdb
        Set dbs = BaseDb
        Set BaseDb = TargetDB
        Set CurrentDB = BaseDb
        MakeAccConnectString
        If ConnectToSql = False Then
          Set BaseDb = dbs
          Set CurrentDB = BaseDb
          Exit Sub
          End If
        Set BaseDb = dbs
        TransferData
End Sub

Private Function TransferData() As Boolean

    Dim k As clsKeyName
    Dim Keyname As String

    On Error GoTo ErrorHandler

    TransferData = False

    '
    ' Start in SQL target.
    ' The SQL target must be empty before conversion.
    ' Do not only check for KeyNames, because a previous failed conversion
    ' may have created tables such as Year without creating KeyNames.
    '
    Set CurrentDB = TargetDB

    If SqlTargetHasUserTables() Then

        If MsgBox("The SQL target database is not empty." & vbCrLf & vbCrLf & _
                  "All existing tables in the SQL target database will be deleted before conversion." & vbCrLf & _
                  "The Access source database will not be changed." & vbCrLf & vbCrLf & _
                  "Do you want to continue?", _
                  vbYesNo + vbExclamation, "NADABAS conversion") = vbNo Then

            Set CurrentDB = BaseDb
            GoTo CleanExit

        End If

        If Not DropAllSqlTargetTables() Then
            Set CurrentDB = BaseDb
            GoTo CleanExit
        End If

    End If

    Load SplashConversionInProgress
    SplashConversionInProgress.Show vbModeless

    SplashConversionInProgress.Label1.Caption = SplashConversionInProgress.lblCreatingBasic.Caption
    DoEvents

    '
    ' Upgrade the source before creating the target. This assigns stable workbook
    ' IDs once, so the same values can be preserved by the conversion.
    '
    Set CurrentDB = BaseDb
    If Not EnsureWorkbookIdentitySchema Then GoTo CleanExit

    '
    ' Create basic NADABAS tables in SQL target.
    '
    Set CurrentDB = TargetDB
    CreateBaseTables
    CreateWorkbookTable

    '
    ' Open source Access database.
    '
    Set CurrentDB = BaseDb
    OpenDb

    '
    ' Copy required base tables.
    ' These should already have been created in target by CreateBaseTables/CreateWorkbookTable.
    '
    If Not CopyTable("WorkbookIdentities") Then GoTo CleanExit
    If Not CopyTable("Workbooks") Then GoTo CleanExit
    If Not CopyTable("Keynames") Then GoTo CleanExit
    If Not CopyTable("Administrators") Then GoTo CleanExit

    '
    ' Copy optional/system tables if they exist in source.
    ' Important: DBTableExists must be checked against BaseDb.
    '

    Set CurrentDB = BaseDb
    If DBTableExists("Permissions") Then
        Set CurrentDB = TargetDB
        CreateTablePermissions
        Set CurrentDB = BaseDb
        If Not CopyTable("Permissions") Then GoTo CleanExit
    End If

    Set CurrentDB = BaseDb
    If DBTableExists("Year") Then
        Set CurrentDB = TargetDB
        CreateTableYears
        Set CurrentDB = BaseDb
        If Not CopyTable("Year") Then GoTo CleanExit
    End If

    Set CurrentDB = BaseDb
    If DBTableExists("BasePath") Then
        Set CurrentDB = TargetDB
        CreateTableBasePath
        Set CurrentDB = BaseDb
        If Not CopyTable("BasePath") Then GoTo CleanExit
    End If

    Set CurrentDB = BaseDb
    If DBTableExists("BackupPath") Then
        Set CurrentDB = TargetDB
        CreateTableBackupPath
        Set CurrentDB = BaseDb
        If Not CopyTable("BackupPath") Then GoTo CleanExit
    End If

    Set CurrentDB = BaseDb
    If DBTableExists("Documents") Then
        Set CurrentDB = TargetDB
        CreateTableDocuments
        Set CurrentDB = BaseDb
        If Not CopyTable("Documents") Then GoTo CleanExit
    End If

    Set CurrentDB = BaseDb
    If DBTableExists("NADABASVersion") Then
        Set CurrentDB = TargetDB
        CreateTableVersion
        Set CurrentDB = BaseDb
        If Not CopyTable("NADABASVersion") Then GoTo CleanExit
    End If

    '
    ' Do not update/save user settings during copy conversion.
    ' We want Access -> SQL to be a pure copy operation.
    '
    ' Usersettings.SaveSettingsToDB

    Set CurrentDB = BaseDb
    If DBTableExists("UserSettings") Then
        Set CurrentDB = TargetDB
        CreateTableUserSettings
        Set CurrentDB = BaseDb
        If Not CopyTable("UserSettings") Then GoTo CleanExit
    End If

    Set CurrentDB = BaseDb
    If DBTableExists("Classifications") Then
        Set CurrentDB = TargetDB
        CreateTableClassifications
        Set CurrentDB = BaseDb
        If Not CopyTable("Classifications") Then GoTo CleanExit
    End If

    Set CurrentDB = BaseDb
    If DBTableExists("Correspondences") Then
        Set CurrentDB = TargetDB
        CreateTableCorrespondence
        Set CurrentDB = BaseDb
        If Not CopyTable("Correspondences") Then GoTo CleanExit
    End If

    Set CurrentDB = BaseDb
    If DBTableExists("DimensionClass") Then
        Set CurrentDB = TargetDB
        CreateTableDimensionClass
        Set CurrentDB = BaseDb
        If Not CopyTable("DimensionClass") Then GoTo CleanExit
    End If

    Set CurrentDB = BaseDb
    If DBTableExists("BatchList") Then
        Set CurrentDB = TargetDB
        CreateTableBatchList
        Set CurrentDB = BaseDb
        If Not CopyTable("BatchList") Then GoTo CleanExit
    End If

    Set CurrentDB = BaseDb
    If DBTableExists("BatchDescription") Then
        Set CurrentDB = TargetDB
        CreateTableBatchDescription
        Set CurrentDB = BaseDb
        If Not CopyTable("BatchDescription") Then GoTo CleanExit
    End If

    Set CurrentDB = BaseDb
    If DBTableExists("Batch2List") Then
        Set CurrentDB = TargetDB
        CreateTableBatch2List
        Set CurrentDB = BaseDb
        If Not CopyTable("Batch2List") Then GoTo CleanExit
    End If

    Set CurrentDB = BaseDb
    If DBTableExists("Batch2Description") Then
        Set CurrentDB = TargetDB
        CreateTableBatch2Description
        Set CurrentDB = BaseDb
        If Not CopyTable("Batch2Description") Then GoTo CleanExit
    End If

    Set CurrentDB = BaseDb
    If DBTableExists("Descriptions") Then
        Set CurrentDB = TargetDB
        CreateTableDescriptions
        Set CurrentDB = BaseDb
        If Not CopyTable("Descriptions") Then GoTo CleanExit
    End If

    Set CurrentDB = BaseDb
    If DBTableExists("DescriptionDimensions") Then
        Set CurrentDB = TargetDB
        CreateTableDescriptionsDimensions
        Set CurrentDB = BaseDb
        If Not CopyTable("DescriptionDimensions") Then GoTo CleanExit
    End If

    Set CurrentDB = BaseDb
    If DBTableExists("DBGlobals") Then
        Set CurrentDB = TargetDB
        CreateTableDBGlobals
        Set CurrentDB = BaseDb
        If Not CopyTable("DBGlobals") Then GoTo CleanExit
    End If

    Set CurrentDB = BaseDb
    If DBTableExists("ClassificationDescriptions") Then
        Set CurrentDB = TargetDB
        CreateTableClassificationDescriptions
        Set CurrentDB = BaseDb
        If Not CopyTable("ClassificationDescriptions") Then GoTo CleanExit
    End If

    '
    ' Now all key families must be established.
    '
    Set CurrentDB = BaseDb
    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions

    For Each k In BaseDb.KeyNames

        Keyname = k.Keyname

        If Not CreateAndCopyTable(Keyname) Then
            GoTo CleanExit
        End If

    Next k

    '
    ' Verify the converted database and record the migration there as well.
    '
    Set CurrentDB = TargetDB
    If Not EnsureWorkbookIdentitySchema Then GoTo CleanExit

    TransferData = True

CleanExit:

    On Error Resume Next

    Set CurrentDB = BaseDb
    CloseDB

    If Not SplashConversionInProgress Is Nothing Then
        SplashConversionInProgress.Hide
        Unload SplashConversionInProgress
    End If

    Set CurrentDB = BaseDb
    DoEvents

    Exit Function

ErrorHandler:

    MsgBox "Error during database conversion:" & vbCrLf & _
           "Error " & err.Number & ": " & err.Description, _
           vbCritical, "NADABAS conversion"

    Debug.Print "ERROR in TransferData"
    Debug.Print "Error: " & err.Number & " - " & err.Description

    Resume CleanExit

End Function


Private Function CreateAndCopyTable(Keyname As String) As Boolean

    Dim xclsFieldNames As Collection
    Dim FN As clsFieldNames
    Dim ValueFN As clsFieldNames
    Dim keyf As clsKeyName
    Dim valuetype As Integer

    On Error GoTo ErrorHandler

    CreateAndCopyTable = False

    SplashConversionInProgress.Label1.Caption = _
        SplashConversionInProgress.lblKeyFam.Caption & " " & Keyname
    DoEvents

    Debug.Print "CreateAndCopyTable started: " & Keyname

    '
    ' Source is Access/BaseDb.
    ' Target is SQL/TargetDB.
    ' This procedure must copy data, not move it.
    ' Therefore: never delete from BaseDb.
    '

    '
    ' Read key family definition from source database.
    '
    Set CurrentDB = BaseDb
    Set keyf = CurrentDB.KeyNames(Keyname)

    Set xclsFieldNames = New Collection
    Set ValueFN = Nothing

    '
    ' Build list of dimension fields.
    '
    ' Important:
    ' - Value is stored separately because CreateNewKeyFam needs valuetype.
    ' - Standard NADABAS data columns are skipped because CreateNewKeyFam
    '   normally adds them itself.
    ' - This avoids duplicate SQL columns such as Comment or DataArea.
    '
    For Each FN In keyf.TableDefinition

        If UCase$(Trim$(FN.name)) = "VALUE" Then

            Set ValueFN = FN
            Debug.Print "  Found Value field: " & FN.name

        ElseIf IsStandardDataColumn(FN.name) Then

            Debug.Print "  Skipped standard data column: " & FN.name

        ElseIf FieldAlreadyInCollection(xclsFieldNames, FN.name) Then

            Debug.Print "  Skipped duplicate dimension field: " & FN.name

        Else

            xclsFieldNames.Add FN
            Debug.Print "  Added dimension field: " & FN.name

        End If

    Next FN

    '
    ' The Value field is required in a NADABAS KeyFamily.
    '
    If ValueFN Is Nothing Then

        MsgBox "Could not find the Value field in the table definition for KeyFamily:" & vbCrLf & _
               Keyname & vbCrLf & vbCrLf & _
               "The SQL table cannot be created safely.", _
               vbCritical, "NADABAS conversion"

        Debug.Print "ERROR: Missing Value field for KeyFamily: " & Keyname

        GoTo CleanExit

    End If

    Debug.Print "KeyFamily: " & Keyname
    Debug.Print "Value SQLType: " & ValueFN.SQLType
    Debug.Print "Value Length: " & ValueFN.Length
    Debug.Print "Number of dimension fields: " & xclsFieldNames.count

    '
    ' Determine value type used by CreateNewKeyFam.
    '
    Select Case ValueFN.SQLType

        Case ADOX.DataTypeEnum.adVarWChar
            valuetype = 3

        Case ADOX.DataTypeEnum.adDouble
            valuetype = 2

        Case ADOX.DataTypeEnum.adSingle
            valuetype = 1

        Case Else

            Select Case ValueFN.Length

                Case 4
                    valuetype = 1

                Case 8
                    valuetype = 2

                Case Else
                    valuetype = 3

            End Select

    End Select

    Debug.Print "ValueType selected: " & valuetype

    '
    ' Delete/create only in target SQL database.
    ' This is the key point that makes the operation COPY, not MOVE.
    '
    Set CurrentDB = TargetDB

    If CurrentDB Is BaseDb Then
        MsgBox "Internal error: attempted to delete KeyFamily from source database:" & vbCrLf & _
               Keyname, vbCritical, "NADABAS conversion"

        Debug.Print "ERROR: CurrentDB is BaseDb before target delete/create."
        GoTo CleanExit
    End If

    If DBTableExists(Keyname) Then
        Debug.Print "Deleting existing target KeyFamily: " & Keyname
        DeleteKeyFam Keyname
    End If

    Debug.Print "Creating target KeyFamily: " & Keyname
    CreateNewKeyFam Keyname, xclsFieldNames, valuetype

    '
    ' Copy data from Access source to SQL target.
    '
    Set CurrentDB = BaseDb

    If Not CopyTable(Keyname) Then
        Debug.Print "ERROR: CopyTable failed for KeyFamily: " & Keyname
        GoTo CleanExit
    End If

    CreateAndCopyTable = True
    Debug.Print "CreateAndCopyTable completed: " & Keyname

CleanExit:

    On Error Resume Next
    Set CurrentDB = BaseDb
    Exit Function

ErrorHandler:

    MsgBox "Error while creating/copying KeyFamily:" & vbCrLf & _
           Keyname & vbCrLf & vbCrLf & _
           "Error " & err.Number & ": " & err.Description, _
           vbCritical, "NADABAS conversion"

    Debug.Print "ERROR in CreateAndCopyTable"
    Debug.Print "KeyFamily: " & Keyname
    Debug.Print "Error: " & err.Number & " - " & err.Description

    Resume CleanExit

End Function

Private Function CopyTable(sTable As String) As Boolean

    Dim f As Variant
    Dim qf As Object

    On Error GoTo ErrorHandler

    CopyTable = False

    Debug.Print "CopyTable started: " & sTable

    If UCase$(sTable) = "WORKBOOKIDENTITIES" Then
        CopyTable = CopyWorkbookIdentityTable()
        Exit Function
    End If

    '
    ' Open source cursor
    '
    Set CurrentDB = BaseDb
    Debug.Print "Opening source cursor: " & CurrentDB.DbDisplayName & " | " & sTable
    CreateGetCursor "Select * from " & InB(sTable)

    '
    ' Open target cursor
    '
    Set CurrentDB = TargetDB
    Debug.Print "Opening target cursor: " & CurrentDB.DbDisplayName & " | " & sTable
    CreateCursor "Select * from " & InB(sTable)

    '
    ' Copy rows
    '
    Set CurrentDB = BaseDb

    Do While CursorGetEoF = False

        Set CurrentDB = BaseDb
        Set qf = GetAllColumnsGet

        Set CurrentDB = TargetDB
        CursorAddNew

        For Each f In qf

            On Error GoTo PutColumnError
            PutColumn f.name, f.value
            On Error GoTo ErrorHandler

        Next f

        CursorUpdate

        Set CurrentDB = BaseDb
        CursorGetMoveNext

        DoEvents

    Loop

    CopyTable = True

CleanExit:

    On Error Resume Next

    Set CurrentDB = TargetDB
    CloseCursor

    Set CurrentDB = BaseDb
    CloseGetCursor

    Set CurrentDB = BaseDb

    Exit Function

PutColumnError:

    MsgBox "Error while copying table:" & vbCrLf & _
           sTable & vbCrLf & vbCrLf & _
           "Column: " & f.name & vbCrLf & _
           "Value: " & NzForDebug(f.value) & vbCrLf & vbCrLf & _
           "Error " & err.Number & ": " & err.Description, _
           vbCritical, "NADABAS conversion"

    Debug.Print "ERROR in CopyTable / PutColumn"
    Debug.Print "Table: " & sTable
    Debug.Print "Column: " & f.name
    Debug.Print "Value: " & NzForDebug(f.value)
    Debug.Print "Error: " & err.Number & " - " & err.Description

    Resume CleanExit

ErrorHandler:

    MsgBox "Error while copying table:" & vbCrLf & _
           sTable & vbCrLf & vbCrLf & _
           "Error " & err.Number & ": " & err.Description, _
           vbCritical, "NADABAS conversion"

    Debug.Print "ERROR in CopyTable"
    Debug.Print "Table: " & sTable
    Debug.Print "Error: " & err.Number & " - " & err.Description

    Resume CleanExit

End Function

Private Function CopyWorkbookIdentityTable() As Boolean

    Dim identityInsertEnabled As Boolean
    Dim rsIdentities As ADODB.Recordset
    Dim WorkbookID As Long
    Dim WorkbookName As String

    On Error GoTo ErrorHandler

    CopyWorkbookIdentityTable = False

    Set CurrentDB = BaseDb
    Set rsIdentities = CurrentDB.DBCnn.Execute( _
        "SELECT [WorkbookID], [WorkbookName] " & _
        "FROM [WorkbookIdentities] ORDER BY [WorkbookID]")

    Set CurrentDB = TargetDB

    If CurrentDB.DBType = Sqlexpress Then
        CurrentDB.DBCnn.Execute "SET IDENTITY_INSERT [WorkbookIdentities] ON"
        identityInsertEnabled = True
    End If

    Do While Not rsIdentities.EOF
        WorkbookID = CLng(rsIdentities.fields("WorkbookID").value)
        WorkbookName = CStr(rsIdentities.fields("WorkbookName").value)

        CurrentDB.DBCnn.Execute _
            "INSERT INTO [WorkbookIdentities] ([WorkbookID], [WorkbookName]) " & _
            "VALUES (" & CStr(WorkbookID) & ", " & InQ(WorkbookName) & ")"

        rsIdentities.MoveNext
        DoEvents
    Loop

    CopyWorkbookIdentityTable = True

CleanExit:
    On Error Resume Next

    If Not rsIdentities Is Nothing Then rsIdentities.Close
    Set rsIdentities = Nothing

    Set CurrentDB = TargetDB
    If identityInsertEnabled Then
        CurrentDB.DBCnn.Execute "SET IDENTITY_INSERT [WorkbookIdentities] OFF"
    End If

    Set CurrentDB = BaseDb
    Exit Function

ErrorHandler:
    MsgBox "Error while preserving workbook IDs during database conversion:" & _
           vbCrLf & vbCrLf & _
           "Error " & err.Number & ": " & err.Description, _
           vbCritical, "NADABAS conversion"

    Debug.Print "ERROR in CopyWorkbookIdentityTable"
    Debug.Print "Error: " & err.Number & " - " & err.Description

    Resume CleanExit

End Function

Private Function NzForDebug(v As Variant) As String

    On Error GoTo ErrorHandler

    If IsNull(v) Then
        NzForDebug = "<Null>"
    ElseIf IsEmpty(v) Then
        NzForDebug = "<Empty>"
    Else
        NzForDebug = CStr(v)
    End If

    Exit Function

ErrorHandler:
    NzForDebug = "<Unable to convert value>"

End Function

Private Function IsStandardDataColumn(ColumnName As String) As Boolean

    Select Case UCase$(Trim$(ColumnName))

        Case "VALUE", _
             "COMMENT", _
             "FORMULA", _
             "USERNAME", _
             "EXCELFILE", _
             "TIMESTAMP", _
             "DATAAREA", _
             "STATUS", _
             "CHANGED", _
             "LOCKED"

            IsStandardDataColumn = True

        Case Else

            IsStandardDataColumn = False

    End Select

End Function

Private Function FieldAlreadyInCollection(fields As Collection, fieldName As String) As Boolean

    Dim FN As clsFieldNames

    FieldAlreadyInCollection = False

    For Each FN In fields

        If UCase$(Trim$(FN.name)) = UCase$(Trim$(fieldName)) Then
            FieldAlreadyInCollection = True
            Exit Function
        End If

    Next FN

End Function

Private Function SqlTargetHasUserTables() As Boolean

    Dim rs As ADODB.Recordset
    Dim sql As String

    On Error GoTo ErrorHandler

    SqlTargetHasUserTables = False

    If CurrentDB.DBType <> Sqlexpress Then Exit Function
    If CurrentDB.DBCnn Is Nothing Then Exit Function

    sql = "SELECT COUNT(*) AS table_count " & _
          "FROM INFORMATION_SCHEMA.TABLES " & _
          "WHERE TABLE_TYPE = 'BASE TABLE'"

    Set rs = CurrentDB.DBCnn.Execute(sql)

    If Not rs.EOF Then
        SqlTargetHasUserTables = (CLng(rs.fields("table_count").value) > 0)
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "ERROR in SqlTargetHasUserTables: " & err.Number & " - " & err.Description
    Resume CleanExit

End Function

Private Function DropAllSqlTargetTables() As Boolean

    Dim sql As String

    On Error GoTo ErrorHandler

    DropAllSqlTargetTables = False

    If CurrentDB.DBType <> Sqlexpress Then
        MsgBox "Internal error: DropAllSqlTargetTables was called while CurrentDB is not SQL Server.", _
               vbCritical, "NADABAS conversion"
        Exit Function
    End If

    If CurrentDB Is BaseDb Then
        MsgBox "Internal error: attempted to drop tables from the source database.", _
               vbCritical, "NADABAS conversion"
        Exit Function
    End If

    sql = "DECLARE @sql NVARCHAR(MAX) = N''; " & _
          "SELECT @sql = @sql + N'DROP TABLE ' + QUOTENAME(TABLE_SCHEMA) + N'.' + QUOTENAME(TABLE_NAME) + N';' " & _
          "FROM INFORMATION_SCHEMA.TABLES " & _
          "WHERE TABLE_TYPE = 'BASE TABLE'; " & _
          "EXEC sp_executesql @sql;"

    CurrentDB.DBCnn.Execute sql

    DropAllSqlTargetTables = True
    Exit Function

ErrorHandler:
    MsgBox "Could not clear the SQL target database." & vbCrLf & vbCrLf & _
           "Error " & err.Number & ": " & err.Description, _
           vbCritical, "NADABAS conversion"

    Debug.Print "ERROR in DropAllSqlTargetTables"
    Debug.Print "Error: " & err.Number & " - " & err.Description

End Function

Private Function GetSourceBasePathForSqlCopy() As String

    On Error GoTo ErrorHandler

    GetSourceBasePathForSqlCopy = ""

    Set CurrentDB = BaseDb

    If CurrentDB.BasefolderInfo Is Nothing Then
        CurrentDB.LoadBasePath
    End If

    If Not CurrentDB.BasefolderInfo Is Nothing Then

        '
        ' If BasePath is linked to the Access database location,
        ' materialise it as the physical folder of the Access file.
        '
        If CurrentDB.BasefolderInfo.BasePathIsDB _
           Or Trim$(CurrentDB.BasefolderInfo.basePath) = "!" _
           Or Trim$(CurrentDB.BasefolderInfo.basePath) = "" Then

            GetSourceBasePathForSqlCopy = DropBackSlash(GetPath(BaseDb.DBFullName))

        Else

            GetSourceBasePathForSqlCopy = DropBackSlash(CurrentDB.BasefolderInfo.basePath)

        End If

    End If

    If Trim$(GetSourceBasePathForSqlCopy) = "" Then
        GetSourceBasePathForSqlCopy = DropBackSlash(GetPath(BaseDb.DBFullName))
    End If

    Debug.Print "Source base path for SQL copy: " & GetSourceBasePathForSqlCopy

    Exit Function

ErrorHandler:

    Debug.Print "ERROR in GetSourceBasePathForSqlCopy"
    Debug.Print "Error: " & err.Number & " - " & err.Description

    On Error Resume Next
    GetSourceBasePathForSqlCopy = DropBackSlash(GetPath(BaseDb.DBFullName))

End Function


Private Function SetTargetSqlBasePath(SourceBasePath As String) As Boolean

    On Error GoTo ErrorHandler

    SetTargetSqlBasePath = False

    If Trim$(SourceBasePath) = "" Then
        GoTo CleanExit
    End If

    Set CurrentDB = TargetDB

    If CurrentDB.DBType <> Sqlexpress Then
        MsgBox "Internal error: target database is not SQL Server when setting base folder.", _
               vbCritical, "NADABAS conversion"
        GoTo CleanExit
    End If

    If CurrentDB.BasefolderInfo Is Nothing Then
        Set CurrentDB.BasefolderInfo = New clsBasePath
    End If

    '
    ' Important:
    ' In SQL, the base folder must be a real file-system path.
    ' It cannot be "!" because SQL Server has no Access-file location.
    '
    CurrentDB.BasefolderInfo.basePath = DropBackSlash(SourceBasePath)
    CurrentDB.BasefolderInfo.BasePathIsDB = False

    saveBasePath

    Set CurrentDB.BasefolderInfo = Nothing
    CurrentDB.LoadBasePath

    Debug.Print "SQL target base path set to: " & CurrentDB.BasefolderInfo.basePath

    SetTargetSqlBasePath = True

CleanExit:

    On Error Resume Next
    Set CurrentDB = BaseDb
    Exit Function

ErrorHandler:

    Debug.Print "ERROR in SetTargetSqlBasePath"
    Debug.Print "SourceBasePath: " & SourceBasePath
    Debug.Print "Error: " & err.Number & " - " & err.Description

    Resume CleanExit

End Function
