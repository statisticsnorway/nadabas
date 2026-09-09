Attribute VB_Name = "DbCreateTables"
Option Explicit
Option Private Module


Dim dt As ADOX.Table
Dim sIndex As String
Dim IndexComma As String

Dim PreparedTableName As String
Dim PreparedIndexName As String
Dim PreparedIndexTableName As String

'
' Common functions
'

Private Function StripBrackets(ByVal s As String) As String

    s = Trim$(s)
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")

    StripBrackets = s

End Function

Public Sub DropIndexIfExists(TableName As String, IndexName As String)

    If Not IndexExists(TableName, IndexName) Then Exit Sub

    DbExecute "DROP INDEX " & SqlBracket(IndexName) & _
              " ON " & SqlBracket(TableName)

End Sub

Public Sub RecreateUniqueIndex(IndexName As String, TableName As String, _
                               ParamArray ColumnNames() As Variant)

    Dim ColumnName As Variant

    DropIndexIfExists TableName, IndexName
    Prepareindex IndexName, TableName

    For Each ColumnName In ColumnNames
        IndexCol CStr(ColumnName)
    Next ColumnName

    AttachIndex

    If Not IndexExists(TableName, IndexName) Then
        err.Raise vbObjectError + 6101, "RecreateUniqueIndex", _
                  "Unable to create index " & TableName & "." & IndexName
    End If

End Sub

Private Function SqlBracket(ByVal s As String) As String

    s = StripBrackets(s)
    SqlBracket = "[" & Replace(s, "]", "]]") & "]"

End Function

Private Function SqlString(ByVal s As String) As String

    SqlString = "N'" & Replace(s, "'", "''") & "'"

End Function

Private Function DbObjectName(ByVal s As String) As String

    Select Case CurrentDB.DBType

        Case Sqlexpress
            DbObjectName = SqlBracket(s)

        Case accdb, mdb
            DbObjectName = StripBrackets(s)

        Case Else
            DbObjectName = StripBrackets(s)

    End Select

End Function

Private Sub PrepareTable(sTable As String)

    PreparedTableName = StripBrackets(sTable)

    Set dt = New ADOX.Table

    Select Case CurrentDB.DBType

        Case Sqlexpress
            dt.name = SqlBracket(PreparedTableName)

        Case accdb, mdb
            dt.name = PreparedTableName

        Case Else
            dt.name = PreparedTableName

    End Select

End Sub

Private Sub AddStrCol(colname As String, ColWidth As Long, allowNull As Boolean)

    Dim cName As String

    cName = DbObjectName(colname)

    dt.Columns.Append cName, ADOX.DataTypeEnum.adVarWChar, ColWidth

    If allowNull Then
        dt.Columns(cName).Attributes = adColNullable
    End If

End Sub

Private Sub AddLongCol(colname As String, allowNull As Boolean)

    Dim cName As String

    cName = DbObjectName(colname)

    dt.Columns.Append cName, ADOX.DataTypeEnum.adInteger

    If allowNull Then
        dt.Columns(cName).Attributes = adColNullable
    End If

End Sub

Private Sub AddSingleCol(colname As String, allowNull As Boolean)

    Dim cName As String

    cName = DbObjectName(colname)

    dt.Columns.Append cName, ADOX.DataTypeEnum.adSingle

    If allowNull Then
        dt.Columns(cName).Attributes = adColNullable
    End If

End Sub

Private Sub AddDoubleCol(colname As String, allowNull As Boolean)

    Dim cName As String

    cName = DbObjectName(colname)

    dt.Columns.Append cName, ADOX.DataTypeEnum.adDouble

    If allowNull Then
        dt.Columns(cName).Attributes = adColNullable
    End If

End Sub

Private Sub AddDateCol(colname As String, allowNull As Boolean)

    Dim cName As String

    cName = DbObjectName(colname)

    Select Case CurrentDB.DBType

        Case Sqlexpress
            dt.Columns.Append cName, ADOX.DataTypeEnum.adDBTimeStamp

        Case accdb, mdb
            dt.Columns.Append cName, ADOX.DataTypeEnum.adDate

        Case Else
            dt.Columns.Append cName, ADOX.DataTypeEnum.adDate

    End Select

    If allowNull Then
        dt.Columns(cName).Attributes = adColNullable
    End If

End Sub

Private Sub AddMemoCol(colname As String, allowNull As Boolean)

    Dim cName As String

    cName = DbObjectName(colname)

    dt.Columns.Append cName, ADOX.DataTypeEnum.adLongVarWChar

    If allowNull Then
        dt.Columns(cName).Attributes = adColNullable
    End If

End Sub

Private Sub AttachTable()

    On Error GoTo ErrorHandler

    '
    ' Use PreparedTableName, not dt.Name.
    ' For SQL Server dt.Name may be "[Year]", while the actual table name is "Year".
    '
    If DBTableExists(PreparedTableName) Then
        Debug.Print "AttachTable skipped, table already exists: " & PreparedTableName
        Exit Sub
    End If

    Select Case CurrentDB.DBType

        Case Sqlexpress, accdb, mdb
            CurrentDB.DBCat.Tables.Append dt

    End Select

    Exit Sub

ErrorHandler:

    MsgBox "Unable to create table:" & vbCrLf & _
           PreparedTableName & vbCrLf & vbCrLf & _
           "Error " & err.Number & ": " & err.Description, _
           vbCritical, "NADABAS"

    Debug.Print "ERROR in AttachTable"
    Debug.Print "Table: " & PreparedTableName
    Debug.Print "dt.Name: " & dt.name
    Debug.Print "Error: " & err.Number & " - " & err.Description

End Sub

Private Sub Prepareindex(IndexName As String, sTable As String)

    PreparedIndexName = StripBrackets(IndexName)
    PreparedIndexTableName = StripBrackets(sTable)

    Select Case CurrentDB.DBType

        Case Sqlexpress
            sIndex = "CREATE UNIQUE NONCLUSTERED INDEX " & _
                     SqlBracket(PreparedIndexName) & _
                     " ON " & SqlBracket(PreparedIndexTableName) & vbCrLf & "("

        Case accdb, mdb
            sIndex = "CREATE UNIQUE INDEX " & _
                     SqlBracket(PreparedIndexName) & _
                     " ON " & SqlBracket(PreparedIndexTableName) & vbCrLf & "("

        Case Else
            sIndex = "CREATE UNIQUE INDEX " & _
                     SqlBracket(PreparedIndexName) & _
                     " ON " & SqlBracket(PreparedIndexTableName) & vbCrLf & "("

    End Select

    IndexComma = " "

End Sub

Private Sub IndexCol(colname As String)

    sIndex = sIndex & IndexComma & SqlBracket(colname) & " ASC"
    IndexComma = ","

End Sub

Private Sub AttachIndex()

    On Error GoTo ErrorHandler

    If IndexExists(PreparedIndexTableName, PreparedIndexName) Then
        Debug.Print "AttachIndex skipped, index already exists: " & _
                    PreparedIndexTableName & "." & PreparedIndexName
        Exit Sub
    End If

    sIndex = sIndex & ")"
    IndexComma = ", "

    CurrentDB.DBCnn.Execute sIndex

    Exit Sub

ErrorHandler:

    MsgBox "Unable to create index:" & vbCrLf & _
           PreparedIndexTableName & "." & PreparedIndexName & vbCrLf & vbCrLf & _
           "Error " & err.Number & ": " & err.Description, _
           vbCritical, "NADABAS"

    Debug.Print "ERROR in AttachIndex"
    Debug.Print "Table: " & PreparedIndexTableName
    Debug.Print "Index: " & PreparedIndexName
    Debug.Print "SQL: " & sIndex
    Debug.Print "Error: " & err.Number & " - " & err.Description

End Sub

Private Function IndexExists(TableName As String, IndexName As String) As Boolean

    Dim rs As ADODB.Recordset
    Dim ix As ADOX.Index
    Dim sql As String

    On Error GoTo ErrorHandler

    IndexExists = False

    TableName = StripBrackets(TableName)
    IndexName = StripBrackets(IndexName)

    Select Case CurrentDB.DBType

        Case Sqlexpress

            sql = "SELECT COUNT(*) AS n " & _
                  "FROM sys.indexes " & _
                  "WHERE object_id = OBJECT_ID(" & SqlString("dbo." & TableName) & ") " & _
                  "AND name = " & SqlString(IndexName)

            Set rs = CurrentDB.DBCnn.Execute(sql)

            If Not rs.EOF Then
                IndexExists = (CLng(rs.fields("n").value) > 0)
            End If

        Case accdb, mdb

            CurrentDB.DBCat.Tables.Refresh
            CurrentDB.DBCat.Tables(TableName).Indexes.Refresh

            For Each ix In CurrentDB.DBCat.Tables(TableName).Indexes

                If UCase$(Trim$(ix.name)) = UCase$(Trim$(IndexName)) Then
                    IndexExists = True
                    Exit Function
                End If

            Next ix

    End Select

CleanExit:

    On Error Resume Next

    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing

    Exit Function

ErrorHandler:

    Debug.Print "ERROR in IndexExists"
    Debug.Print "Table: " & TableName
    Debug.Print "Index: " & IndexName
    Debug.Print "Error: " & err.Number & " - " & err.Description

    Resume CleanExit

End Function

Private Sub CreateView(VName As String, ssql As String)

    Dim s As String

    On Error GoTo ErrorHandler

    If ViewExists(VName) Then
        Debug.Print "CreateView skipped, view already exists: " & VName
        Exit Sub
    End If

    Select Case CurrentDB.DBType

        Case Sqlexpress, accdb, mdb
            s = "CREATE VIEW " & SqlBracket(VName) & " AS " & ssql
            CurrentDB.DBCnn.Execute s

    End Select

    Exit Sub

ErrorHandler:

    MsgBox "Unable to create view:" & vbCrLf & _
           VName & vbCrLf & vbCrLf & _
           "Error " & err.Number & ": " & err.Description, _
           vbCritical, "NADABAS"

    Debug.Print "ERROR in CreateView"
    Debug.Print "View: " & VName
    Debug.Print "SQL: " & s
    Debug.Print "Error: " & err.Number & " - " & err.Description

End Sub

Private Function ViewExists(VName As String) As Boolean

    Dim rs As ADODB.Recordset
    Dim sql As String

    On Error GoTo ErrorHandler

    ViewExists = False
    VName = StripBrackets(VName)

    Select Case CurrentDB.DBType

        Case Sqlexpress

            sql = "SELECT COUNT(*) AS n " & _
                  "FROM INFORMATION_SCHEMA.VIEWS " & _
                  "WHERE TABLE_NAME = " & SqlString(VName)

            Set rs = CurrentDB.DBCnn.Execute(sql)

            If Not rs.EOF Then
                ViewExists = (CLng(rs.fields("n").value) > 0)
            End If

        Case accdb, mdb

            '
            ' Access view/query handling varies in this codebase.
            ' Keep conservative behaviour.
            '
            ViewExists = False

    End Select

CleanExit:

    On Error Resume Next

    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing

    Exit Function

ErrorHandler:

    Debug.Print "ERROR in ViewExists"
    Debug.Print "View: " & VName
    Debug.Print "Error: " & err.Number & " - " & err.Description

    Resume CleanExit

End Function




'
' Functions that creates new tables and views in the database
'

Public Sub CreateBaseTables()



' Database is open and is not a nadabas db

'
'   drop any tables that may exist
'
      DropTable "KeyNames"
      DropTable "Administrators"
'
'  now create the tables
'

      PrepareTable "KeyNames"
      AddStrCol "KeyName", 50, False
      AttachTable

     Prepareindex "PrimaryIndex", "KeyNames"
     IndexCol "KeyName"
     AttachIndex

      If CurrentDB.DbIsExch Then Exit Sub

      PrepareTable "Administrators"
      AddStrCol "UserName", 50, False
      AttachTable

      Prepareindex "PrimaryIndex", "Administrators"
      IndexCol "UserName"
      AttachIndex




 End Sub

Public Sub CreateTableWorkbookIdentities()

    If DBTableExists("WorkbookIdentities") Then Exit Sub

    Select Case CurrentDB.DBType
        Case Sqlexpress
            DbExecute "CREATE TABLE [WorkbookIdentities] (" & _
                      "[WorkbookID] INT IDENTITY(1,1) NOT NULL, " & _
                      "[WorkbookName] NVARCHAR(" & WORKBOOK_NAME_MAX_LENGTH & ") NOT NULL)"
        Case accdb, mdb
            DbExecute "CREATE TABLE [WorkbookIdentities] (" & _
                      "[WorkbookID] AUTOINCREMENT, " & _
                      "[WorkbookName] TEXT(" & WORKBOOK_NAME_MAX_LENGTH & ") NOT NULL)"
    End Select

    Prepareindex "PrimaryIndex", "WorkbookIdentities"
    IndexCol "WorkbookID"
    AttachIndex

    Prepareindex "WorkbookNameIndex", "WorkbookIdentities"
    IndexCol "WorkbookName"
    AttachIndex

End Sub

Public Sub CreateTableSchemaMigrations()

    If DBTableExists("NADABASSchemaMigrations") Then Exit Sub

    PrepareTable "NADABASSchemaMigrations"
    AddStrCol "MigrationName", 100, False
    AddDateCol "AppliedAt", False
    AttachTable

    Prepareindex "PrimaryIndex", "NADABASSchemaMigrations"
    IndexCol "MigrationName"
    AttachIndex

End Sub

Public Sub CreateWorkbookTable()
      DropTable "Workbooks"

      CreateTableWorkbookIdentities
      PrepareTable "Workbooks"
      AddLongCol "WorkbookID", False
      AddStrCol "WorkBookName", WORKBOOK_NAME_MAX_LENGTH, False
      AddStrCol "Title", 100, True
      AddStrCol "GroupName", 50, True
      AddStrCol "Path", 255, True
      AddStrCol "ReservedBy", 50, True
      AddStrCol "Status", 50, True
      AddDateCol "ReservedDate", True
      AddLongCol "isdirty", True
      AddDateCol "LastGet", True
      AddDateCol "LastPut", True
      AddLongCol "Updateorder", True
      AttachTable

      Prepareindex "PrimaryIndex", "Workbooks"
      IndexCol "WorkbookID"
      AttachIndex

      Prepareindex "WorkbookNameIndex", "Workbooks"
      IndexCol "WorkBookName"
      AttachIndex

End Sub

Public Sub CreateTablePermissions()

      PrepareTable "Permissions"
      AddLongCol "WorkbookID", False
      AddStrCol "WorkBookName", WORKBOOK_NAME_MAX_LENGTH, False
      AddStrCol "User", 100, True
      AttachTable

End Sub


Public Sub CreateTableYears()

      PrepareTable "Year"
      AddStrCol "Name", 8, False
      AddStrCol "StartYear", 8, True
      AddStrCol "EndYear", 8, True
      AddStrCol "DefaultFormat", 12, True
      AttachTable

End Sub


Public Sub CreateTableBasePath()
      PrepareTable "BasePath"
      AddStrCol "BasePath", 255, False
      AttachTable
End Sub

Public Sub CreateTableBackupPath()
      PrepareTable "BackupPath"
      AddStrCol "BackupPath", 255, False
      AttachTable
End Sub

Public Sub CreateTableDocuments()

      PrepareTable "Documents"
      AddStrCol "Name", 50, False
      AddStrCol "Path", 255, False
      AddLongCol "Level", True
      AddStrCol "DGroup", 50, True
      AddLongCol "WorkbookID", True
      AddStrCol "Workbook", WORKBOOK_NAME_MAX_LENGTH, True
      AttachTable

End Sub


Public Sub CreateTableVersion()

      PrepareTable "NADABASVersion"
      AddStrCol "VersionNumber", 12, False
      AttachTable

End Sub
Public Sub CreateTableUserSettings()

      PrepareTable "UserSettings"
      AddLongCol "LoadedCellBackCol", False
      AddLongCol "SavedCellBackCol", False
      AddLongCol "LoadedCellTextCol", False

      AddLongCol "SavedCellTextCol", False
      AddLongCol "LoadedCellFont", False
      AddLongCol "SavedCellFont", False
      AddLongCol "WasLoadedBackCol", False
      AddLongCol "MissedCellBackCol", False
      AddLongCol "DBLinksCol", False
      AddLongCol "TableDefCol", False
      AddLongCol "DefAreaCol", False
      AddLongCol "DescriptionsCol", False
      AddStrCol "Password", 20, True
      AddLongCol "PasswordOnWB", True
      AddLongCol "BooksInBase", True
      AddLongCol "DocsInBase", True
      AddLongCol "DropTestForChange", True
      AddLongCol "AllowUnregistered", True
      AddLongCol "DoNotSaveEmptyCells", True
      AddLongCol "UpdateExternalLinks", True
      AddLongCol "NoTestAtOpen", True
      AddLongCol "NoTestAtClose", True
      AddLongCol "UsersMayRunBatch", True
      AddStrCol "SepChar", 1, False
      AddLongCol "SaveAfterLoad", True
      AddLongCol "SaveAfterSave", True
      AddLongCol "Dropbox", True
      AddLongCol "SatelliteSystem", True
      AddLongCol "ImportExport", True
      AddLongCol "CheckForUpdates", True
      AttachTable

End Sub

Public Sub DeleteKeyFam(Keyname As String)
     DropTable (Keyname)
End Sub

Public Function CreateNewKeyFam(FName As String, Dimensions As Collection, valuetype As Integer) As Boolean

Dim s As String
Dim sand As String
Dim f As field
Dim x As clsFieldNames

 '    On Error GoTo someerror

     PrepareTable FName
     For Each x In Dimensions
        AddStrCol x.name, x.Length, False
     Next x
     Select Case valuetype
     Case 1
        AddSingleCol "Value", True
     Case 2
        AddDoubleCol "Value", True
     Case 3
        AddStrCol "Value", 255, True
     End Select
     AddMemoCol "Comment", True
     AddStrCol "Formula", 255, True
     AddStrCol "Username", 50, True
     AddStrCol "ExcelFile", 255, True
     AddDateCol "Timestamp", True
     AddStrCol "DataArea", 50, True
     AttachTable

     Prepareindex "PrimaryIndex", FName
     For Each x In Dimensions
         IndexCol InB(x.name)
     Next x
     AttachIndex



    CreateNewKeyFam = True
    Exit Function
someerror:
    MsgBox err.Description, vbOKOnly

    CreateNewKeyFam = False
    Exit Function
End Function




Public Sub CreateTableClassifications()

      PrepareTable "Classifications"
      AddStrCol "ClassName", 50, False
      AddStrCol "Code", 50, False
      AddStrCol "Title", 255, True
      AttachTable

     Prepareindex "PrimaryIndex", "Classifications"
     IndexCol "ClassName"
     IndexCol "Code"
     AttachIndex
End Sub

Public Sub CreateTableCorrespondence()

      PrepareTable "Correspondences"
      AddStrCol "Fromclass", 50, False
      AddStrCol "FromCode", 50, False
      AddStrCol "Toclass", 50, False
      AddStrCol "ToCode", 50, False
      AttachTable

     Prepareindex "PrimaryIndex", "Correspondences"
     IndexCol "Fromclass"
     IndexCol "FromCode"
     IndexCol "Toclass"
     AttachIndex
End Sub

Public Sub CreateTableDimensionClass()

      PrepareTable "DimensionClass"
      AddStrCol "DimensionName", 50, False
      AddStrCol "ClassName", 50, False
      AttachTable

End Sub
Public Sub CreateTableBatchList()

      PrepareTable "BatchList"
      AddStrCol "Listname", 50, False
      AddLongCol "ItemNo", False
      AddLongCol "WorkbookID", False
      AddStrCol "Workbookname", WORKBOOK_NAME_MAX_LENGTH, False
      AttachTable

     Prepareindex "PrimaryIndex", "BatchList"
     IndexCol "Listname"
     IndexCol "ItemNo"
     AttachIndex

End Sub

Public Sub CreateTableBatchDescription()

      PrepareTable "BatchDescription"
      AddStrCol "Listname", 50, False
      AddMemoCol "Description", True
      AttachTable

     Prepareindex "PrimaryIndex", "BatchDescription"
     IndexCol "Listname"
     AttachIndex
End Sub

Public Sub CreateTableBatch2List()

      PrepareTable "Batch2List"
      AddStrCol "Listname", 50, False
      AddLongCol "ItemNo", False
      AddStrCol "Batchname", 50, False
      AttachTable

     Prepareindex "PrimaryIndex", "Batch2List"
     IndexCol "Listname"
     IndexCol "ItemNo"
     AttachIndex

End Sub

Public Sub CreateTableBatch2Description()

      PrepareTable "Batch2Description"
      AddStrCol "Listname", 50, False
      AddMemoCol "Description", True
      AttachTable

     Prepareindex "PrimaryIndex", "Batch2Description"
     IndexCol "Listname"
     AttachIndex
End Sub


Public Sub CreateTableDescriptions()

      PrepareTable "Descriptions"
      AddLongCol "WorkbookID", False
      AddStrCol "WorkbookName", WORKBOOK_NAME_MAX_LENGTH, False
      AddStrCol "DataAreaName", 50, False
      AddStrCol "TableName", 50, False
      AddStrCol "GetPut", 10, False
      AddMemoCol "Description", True
      AttachTable

     Prepareindex "PrimaryIndex", "Descriptions"
     IndexCol "WorkbookID"
     IndexCol "DataAreaName"
     AttachIndex
 End Sub
 Public Sub CreateTableDescriptionsDimensions()
     PrepareTable "DescriptionDimensions"
     AddLongCol "WorkbookID", False
     AddStrCol "WorkbookName", WORKBOOK_NAME_MAX_LENGTH, False
     AddStrCol "DataAreaName", 50, False
     AddLongCol "DimensionNumber", False
     AddStrCol "DimensionName", 50, False
     AddStrCol "Classification", 50, True
     AddStrCol "ConstantValue", 50, True
     AttachTable

     Prepareindex "PrimaryIndex", "DescriptionDimensions"
     IndexCol "WorkbookID"
     IndexCol "DataAreaName"
     IndexCol "DimensionNumber"
     AttachIndex


End Sub

Public Sub CreateTableDBGlobals()

      PrepareTable "DBGlobals"
      AddLongCol "IDNum", False
      AddStrCol "Name", 50, False
      AddStrCol "Value", 50, False
      AttachTable

     Prepareindex "PrimaryIndex", "DBGlobals"
     IndexCol "IDNum"
     AttachIndex


    End Sub

Public Sub CreateTableClassificationDescriptions()

      PrepareTable "ClassificationDescriptions"
      AddStrCol "ClassName", 50, False
      AddMemoCol "Description", True
      AttachTable

     Prepareindex "PrimaryIndex", "ClassificationDescriptions"
     IndexCol "ClassName"
     AttachIndex

    End Sub

Public Sub CreateTableDataLinks()

    If DBTableExists("DataLinks") Then
        Debug.Print "CreateTableDataLinks skipped: DataLinks already exists"
        Exit Sub
    End If

    PrepareTable "DataLinks"
    AddStrCol "KeyFamily", 50, False
    AddLongCol "TargetWorkbookID", False
    AddStrCol "TargetWB", WORKBOOK_NAME_MAX_LENGTH, False
    AddStrCol "TargetDataArea", 50, False
    AddLongCol "SourceWorkbookID", False
    AddStrCol "SourceWB", WORKBOOK_NAME_MAX_LENGTH, False
    AddStrCol "SourceDataArea", 50, False

    AttachTable

    Prepareindex "PrimaryIndex", "DataLinks"
    IndexCol "KeyFamily"
    IndexCol "TargetWorkbookID"
    IndexCol "TargetDataArea"
    IndexCol "SourceWorkbookID"
    IndexCol "SourceDataArea"
    AttachIndex

End Sub


Public Sub DeleteAllTables()
'
' only used when reconverting a database.
'
'
Dim Keyname As clsKeyName


    If MsgBox(GetMsg("M149"), vbYesNo) = vbNo Then 'Are you sure you want to completly renew the database
       Exit Sub
    End If
    OpenDb
    CurrentDB.LoadKeyNames

 For Each Keyname In CurrentDB.KeyNames
     DropTable (Keyname.Keyname)
 Next Keyname


    DropTable "Workbooks"
    DropTable "Keynames"
    DropTable "Administrators"
    DropTable "Permissions"
    DropTable "Year"
    DropTable "BasePath"
    DropTable "BackupPath"
    DropTable "Documents"
    DropTable "NADABASVersion"
    DropTable "UserSettings"
    DropTable "Classifications"
    DropTable "Correspondences"
    DropTable "DimensionClass"
    DropTable "BatchList"
    DropTable "BatchDescription"
    DropTable "Batch2List"
    DropTable "Batch2Description"
    DropTable "Descriptions"
    DropTable "DescriptionDimensions"
    DropTable "DBGlobals"
    DropTable "ClassificationDescriptions"
    DropTable "DataLinks"
    DropTable "NADABASSchemaMigrations"
    DropTable "WorkbookIdentities"
    CloseDB
End Sub


Public Sub AlterFieldLen(sTable As String, Dimensionname As String, newlen As Integer)

    Select Case CurrentDB.DBType
      Case Sqlexpress
            DbExecute "Alter TABLE " & InB(sTable) & " ALTER COLUMN " & InB(Dimensionname) & " nvarchar(" & newlen & ")"
      Case accdb, mdb
            DbExecute "Alter TABLE " & InB(sTable) & " ALTER COLUMN " & InB(Dimensionname) & " TEXT(" & newlen & ")"
    End Select

End Sub

Public Sub DropTable(sTable As String)

   On Error Resume Next
    Select Case CurrentDB.DBType
      Case Sqlexpress
        DbExecute "Drop Table " & sTable
      Case accdb, mdb
        DbExecute "Drop Table [" & sTable & "]"
    End Select

End Sub


Public Function DBTableExists(sTablename As String) As Boolean

    Dim ob As Object
    Dim rs As ADODB.Recordset
    Dim sql As String

    On Error GoTo ErrorHandler

    DBTableExists = False

    sTablename = StripBrackets(sTablename)

    Select Case CurrentDB.DBType

        Case Sqlexpress

            If CurrentDB.DBCnn Is Nothing Then Exit Function

            sql = "SELECT COUNT(*) AS n " & _
                  "FROM INFORMATION_SCHEMA.TABLES " & _
                  "WHERE TABLE_TYPE = 'BASE TABLE' " & _
                  "AND TABLE_NAME = " & SqlString(sTablename)

            Set rs = CurrentDB.DBCnn.Execute(sql)

            If Not rs.EOF Then
                DBTableExists = (CLng(rs.fields("n").value) > 0)
            End If

        Case accdb, mdb

            On Error Resume Next
            CurrentDB.DBCat.Tables.Refresh
            Set ob = CurrentDB.DBCat.Tables(sTablename)
            DBTableExists = Not ob Is Nothing
            On Error GoTo 0

        Case Else

            DBTableExists = False

    End Select

CleanExit:

    CloseRecordsetSafely rs
    Set ob = Nothing

    Exit Function

ErrorHandler:

    Debug.Print "ERROR in DBTableExists"
    Debug.Print "Table: " & sTablename
    Debug.Print "Error: " & err.Number & " - " & err.Description

    Resume CleanExit

End Function

Public Function DBColumnExists(TableName As String, ColumnName As String) As Boolean

    Dim ob As Object
    Dim rs As ADODB.Recordset
    Dim sql As String

    On Error GoTo CleanExit

    TableName = StripBrackets(TableName)
    ColumnName = StripBrackets(ColumnName)

    Select Case CurrentDB.DBType
        Case Sqlexpress
            sql = "SELECT COUNT(*) AS n FROM INFORMATION_SCHEMA.COLUMNS " & _
                  "WHERE TABLE_NAME = " & SqlString(TableName) & _
                  " AND COLUMN_NAME = " & SqlString(ColumnName)
            Set rs = CurrentDB.DBCnn.Execute(sql)
            If Not rs.EOF Then
                DBColumnExists = (CLng(rs.fields("n").value) > 0)
            End If
        Case accdb, mdb
            On Error Resume Next
            CurrentDB.DBCat.Tables.Refresh
            CurrentDB.DBCat.Tables(TableName).Columns.Refresh
            Set ob = CurrentDB.DBCat.Tables(TableName).Columns(ColumnName)
            DBColumnExists = Not ob Is Nothing
            On Error GoTo 0
        Case Else
            DBColumnExists = False
    End Select

CleanExit:
    CloseRecordsetSafely rs
    Set ob = Nothing

End Function

Public Function DBColumnSize(TableName As String, ColumnName As String) As Long

    Dim rs As ADODB.Recordset
    Dim sql As String

    On Error GoTo CleanExit

    TableName = StripBrackets(TableName)
    ColumnName = StripBrackets(ColumnName)

    Select Case CurrentDB.DBType
        Case Sqlexpress
            sql = "SELECT CHARACTER_MAXIMUM_LENGTH AS n " & _
                  "FROM INFORMATION_SCHEMA.COLUMNS " & _
                  "WHERE TABLE_NAME = " & SqlString(TableName) & _
                  " AND COLUMN_NAME = " & SqlString(ColumnName)
            Set rs = CurrentDB.DBCnn.Execute(sql)
            If Not rs.EOF Then
                If Not IsNull(rs.fields("n").value) Then
                    DBColumnSize = CLng(rs.fields("n").value)
                End If
            End If
        Case accdb, mdb
            CurrentDB.DBCat.Tables.Refresh
            CurrentDB.DBCat.Tables(TableName).Columns.Refresh
            DBColumnSize = CurrentDB.DBCat.Tables(TableName).Columns(ColumnName).DefinedSize
        Case Else
            DBColumnSize = 0
    End Select

CleanExit:
    CloseRecordsetSafely rs

End Function

Private Sub CloseRecordsetSafely(ByRef rs As ADODB.Recordset)

    On Error Resume Next

    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing

End Sub
