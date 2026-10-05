Attribute VB_Name = "cmdKeyFamily"
Option Explicit
Option Private Module
'
' command related to key families
'

Public Sub ShowKeyFamily()
' *******************
' called from Ribbon
' *******************

'   *********************************************
'   *                                           *
'   * Show Key Families                         *
'   *                                           *
'   *********************************************
Dim keyf As clsKeyName

     CurrentDB.LoadKeyNames
     CurrentDB.LoadDimensions
     For Each keyf In CurrentDB.KeyNames
        If Not DBTableExists(keyf.Keyname) Then
            MsgBox GetMsg1("M214", keyf.Keyname), vbCritical
        Exit Sub
        End If
     Next keyf

     Load frmKeyFamily
     frmKeyFamily.Initialize False
     frmKeyFamily.Show vbModal
     Unload frmKeyFamily

     CurrentDB.DimensionClassesIsLoaded = False  ' just in case
End Sub

'   *********************************************
'   *                                           *
'   *  Create Key Fmiliy                        *
'   *                                           *
'   *********************************************

Public Sub CreateKeyFamilyExch()
' *******************
' called from Ribbon
' *******************

    Set CurrentDB = ExchDB
    CreateKeyFamily
    Set CurrentDB = BaseDb
End Sub

'
' Code related to creation and maintenance of Key Families
'
Public Sub CreateKeyFamily()
' *******************
' called from Ribbon
' *******************

' Create a keyfamily in the CURRENT databAse
' Use dlgCreateKeyFamily to get inof needed
'
Dim b As Boolean
     CurrentDB.LoadKeyNames
     CurrentDB.LoadDimensions      'set dimensions, tabledefinitions in currentDB

     Load frmCreateKeyFamily
     frmCreateKeyFamily.Initialize
     frmCreateKeyFamily.Show vbModal
     Unload frmCreateKeyFamily
End Sub


'   *********************************************
'   *                                           *
'   *  Manage Key Families                      *
'   *                                           *
'   *********************************************

Public Sub ManageKeyFamiliesExch()
' *******************
' called from Ribbon
' *******************
     Set CurrentDB = ExchDB
     ManageKeyFamilies
     Set CurrentDB = BaseDb
End Sub


Public Sub ManageKeyFamilies()
' *******************
' called from Ribbon
' *******************
Dim keyf As clsKeyName

     CurrentDB.LoadKeyNames
     CurrentDB.LoadDimensions
     For Each keyf In CurrentDB.KeyNames
        If Not DBTableExists(keyf.Keyname) Then
            MsgBox GetMsg1("M214", keyf.Keyname), vbCritical
        Exit Sub
        End If
     Next keyf

     Load frmKeyFamily
     frmKeyFamily.Initialize True
     frmKeyFamily.Show vbModal
     Unload frmKeyFamily
     CurrentDB.DimensionClassesIsLoaded = False  ' just in case
End Sub


Public Sub KeyFamilyDelete(KeyFam As String)
Dim s As String

    OpenDb
    CurrentDB.DeleteKeyName KeyFam
    DropTable KeyFam
    CloseDB
    CurrentDB.DimensionsIsLoaded = False
    CurrentDB.KeynamesIsLoaded = False
    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions

End Sub

'   *********************************************
'   *                                           *
'   *  Structure Editor Helper Routines          *
'   *                                           *
'   *********************************************

Public Sub KeyFamilyGetColumns(Keyname As String, Names As Collection, varlen As Collection, valuetype As Integer)
    Dim keyf As clsKeyName
    Dim field As clsFieldNames

    Set Names = New Collection
    Set varlen = New Collection
    valuetype = 2

    Set keyf = CurrentDB.GetKeyName(Keyname)
    If keyf Is Nothing Then Exit Sub
    If keyf.TableDefinition.count = 0 Then Exit Sub
    For Each field In keyf.TableDefinition
        If UCase(field.name) = "VALUE" Then
            valuetype = field.SQLType
            Exit For
        End If
        Names.Add field.name
        varlen.Add field.Length
    Next field
End Sub

Public Function KeyFamilyRename(OldName As String, NewName As String) As Boolean
    Dim ssql As String
    Dim tbl As Object
    Dim keyf As clsKeyName
    Dim field As clsFieldNames
    Dim idxCols As String
    Dim renamed As Boolean
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo errHandler
    KeyFamilyRename = False

    NewName = Trim(NewName)
    If NewName = "" Then
        MsgBox "Please enter a new Key Family name.", vbExclamation, "NADABAS"
        Exit Function
    End If

    If UCase(OldName) = UCase(NewName) Then
        Exit Function
    End If

    If Not TestValidname(NewName, "Key Family") Then
        Exit Function
    End If

    If CurrentDB.KeyNames.count > 0 Then
        On Error Resume Next
        Set tbl = CurrentDB.KeyNames(NewName)
        On Error GoTo errHandler
        If Not tbl Is Nothing Then
            MsgBox "A Key Family named '" & NewName & "' already exists.", vbCritical, "NADABAS"
            Exit Function
        End If
    End If

    OpenDb

    renamed = False
    Select Case CurrentDB.DBType
        Case accdb, mdb
            On Error Resume Next
            If Not CurrentDB.DBCat Is Nothing Then
                CurrentDB.DBCat.Tables.Refresh
                Set tbl = CurrentDB.DBCat.Tables(OldName)
                If Not tbl Is Nothing Then
                    tbl.name = NewName
                    renamed = True
                End If
            End If
            On Error GoTo errHandler

            ' Fallback if ADOX rename failed
            If Not renamed Then
                DbExecute "SELECT * INTO " & InB(NewName) & " FROM " & InB(OldName)
                idxCols = ""
                Set keyf = CurrentDB.GetKeyName(OldName)
                If Not keyf Is Nothing Then
                    For Each field In keyf.TableDefinition
                        If UCase(field.name) = "VALUE" Then Exit For
                        If idxCols <> "" Then idxCols = idxCols & ", "
                        idxCols = idxCols & InB(field.name)
                    Next field
                End If
                If idxCols <> "" Then
                    DbExecute "CREATE UNIQUE INDEX [PrimaryIndex] ON " & InB(NewName) & " (" & idxCols & ") WITH PRIMARY"
                End If
                DropTable OldName
                renamed = True
            End If

        Case Sqlexpress
            DbExecute "EXEC sp_rename " & InQ(OldName) & ", " & InQ(NewName)
            renamed = True
    End Select

    DbExecute "UPDATE [KeyNames] SET [KeyName] = " & InQ(NewName) & " WHERE [KeyName] = " & InQ(OldName)

    If DBTableExists("DataLinks") Then
        DbExecute "UPDATE [DataLinks] SET [KeyFamily] = " & InQ(NewName) & " WHERE [KeyFamily] = " & InQ(OldName)
    End If

    If DBTableExists("Descriptions") Then
        DbExecute "UPDATE [Descriptions] SET [TableName] = " & InQ(NewName) & " WHERE [TableName] = " & InQ(OldName)
    End If

    On Error Resume Next
    If Not CurrentDB.DBCat Is Nothing Then
        CurrentDB.DBCat.Tables.Refresh
    End If
    On Error GoTo errHandler

    CloseDB

    CurrentDB.KeynamesIsLoaded = False
    CurrentDB.DimensionsIsLoaded = False
    CurrentDB.DimensionClassesIsLoaded = False
    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions

    MsgBox "Key Family '" & OldName & "' was successfully renamed to '" & NewName & "'." & vbCrLf & vbCrLf & _
           "Note: If any Workbooks reference '" & OldName & "' in their DBLinks / DBDef sheets, " & _
           "please update the table name in those sheets.", vbInformation, "NADABAS"

    KeyFamilyRename = True
    Exit Function

errHandler:
    errNum = err.Number
    errDesc = err.Description
    CloseDB
    ShowNadabasError "KeyFamilyRename", errNum, errDesc
    KeyFamilyRename = False
End Function

Public Function KeyFamilyAddDimension(KeyFam As String, DimName As String, DimLen As Long, DefaultVal As String) As Boolean
    Dim keyf As clsKeyName
    Dim field As clsFieldNames
    Dim existingDims As Collection
    Dim rowCount As Long
    Dim idxCols As String
    Dim colDefs As String
    Dim insCols As String
    Dim selCols As String
    Dim valType As Integer
    Dim valColDef As String
    Dim stgTable As String
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo errHandler
    KeyFamilyAddDimension = False

    DimName = Trim(DimName)
    If DimName = "" Then
        MsgBox "Please enter a dimension name.", vbExclamation, "NADABAS"
        Exit Function
    End If

    If Not TestValidname(DimName, "Dimension") Then
        Exit Function
    End If

    If DimLen <= 0 Or DimLen > 50 Then
        MsgBox "Dimension length must be between 1 and 50.", vbExclamation, "NADABAS"
        Exit Function
    End If

    Select Case UCase(DimName)
        Case "VALUE", "COMMENT", "FORMULA", "USERNAME", "EXCELFILE", "TIMESTAMP", "DATAAREA"
            MsgBox "'" & DimName & "' is a reserved column name in NADABAS.", vbCritical, "NADABAS"
            Exit Function
    End Select

    OpenDb

    Set existingDims = New Collection
    Set keyf = CurrentDB.GetKeyName(KeyFam)
    valType = 2

    If Not keyf Is Nothing Then
        For Each field In keyf.TableDefinition
            If UCase(field.name) = "VALUE" Then
                valType = field.SQLType
                Exit For
            End If
            If UCase(field.name) = UCase(DimName) Then
                CloseDB
                MsgBox "Dimension '" & DimName & "' already exists in Key Family '" & KeyFam & "'.", vbExclamation, "NADABAS"
                Exit Function
            End If
            existingDims.Add field
        Next field
    End If

    If existingDims.count = 0 Then
        CloseDB
        MsgBox "Cannot find existing dimensions for Key Family '" & KeyFam & "'.", vbCritical, "NADABAS"
        Exit Function
    End If

    rowCount = 0
    If CreateCursor("SELECT COUNT(*) AS Cnt FROM " & InB(KeyFam)) Then
        If Not CursorEoF Then
            rowCount = CLng(GetColumn("Cnt"))
        End If
        CloseCursor
    End If

    If rowCount > 0 And Trim(DefaultVal) = "" Then
        CloseDB
        MsgBox "This Key Family contains " & rowCount & " data records." & vbCrLf & _
               "You must provide a default value for existing records because primary key dimensions cannot be empty.", _
               vbExclamation, "NADABAS"
        Exit Function
    End If

    Select Case valType
        Case ADOX.DataTypeEnum.adSingle:
            valColDef = "SINGLE"
        Case ADOX.DataTypeEnum.adDouble:
            valColDef = "DOUBLE"
        Case ADOX.DataTypeEnum.adVarWChar:
            valColDef = IIf(CurrentDB.DBType = Sqlexpress, "NVARCHAR(255)", "TEXT(255)")
        Case Else
            valColDef = "DOUBLE"
    End Select

    colDefs = ""
    idxCols = ""
    insCols = ""
    selCols = ""

    For Each field In existingDims
        If colDefs <> "" Then colDefs = colDefs & ", "
        If idxCols <> "" Then idxCols = idxCols & ", "
        If insCols <> "" Then insCols = insCols & ", "
        If selCols <> "" Then selCols = selCols & ", "

        Select Case CurrentDB.DBType
            Case accdb, mdb
                colDefs = colDefs & InB(field.name) & " TEXT(" & field.Length & ") NOT NULL"
            Case Sqlexpress
                colDefs = colDefs & InB(field.name) & " NVARCHAR(" & field.Length & ") NOT NULL"
        End Select

        idxCols = idxCols & InB(field.name)
        insCols = insCols & InB(field.name)
        selCols = selCols & InB(field.name)
    Next field

    ' Add new dimension
    Select Case CurrentDB.DBType
        Case accdb, mdb
            colDefs = colDefs & ", " & InB(DimName) & " TEXT(" & DimLen & ") NOT NULL"
        Case Sqlexpress
            colDefs = colDefs & ", " & InB(DimName) & " NVARCHAR(" & DimLen & ") NOT NULL"
    End Select
    idxCols = idxCols & ", " & InB(DimName)
    insCols = insCols & ", " & InB(DimName)
    selCols = selCols & ", " & InQ(Trim(DefaultVal))

    ' Standard metadata columns
    Select Case CurrentDB.DBType
        Case accdb, mdb
            colDefs = colDefs & ", [Value] " & valColDef & _
                                ", [Comment] MEMO" & _
                                ", [Formula] TEXT(255)" & _
                                ", [Username] TEXT(50)" & _
                                ", [ExcelFile] TEXT(255)" & _
                                ", [Timestamp] DATETIME" & _
                                ", [DataArea] TEXT(50)"
        Case Sqlexpress
            colDefs = colDefs & ", [Value] " & valColDef & _
                                ", [Comment] NVARCHAR(MAX)" & _
                                ", [Formula] NVARCHAR(255)" & _
                                ", [Username] NVARCHAR(50)" & _
                                ", [ExcelFile] NVARCHAR(255)" & _
                                ", [Timestamp] DATETIME" & _
                                ", [DataArea] NVARCHAR(50)"
    End Select

    insCols = insCols & ", [Value], [Comment], [Formula], [Username], [ExcelFile], [Timestamp], [DataArea]"
    selCols = selCols & ", [Value], [Comment], [Formula], [Username], [ExcelFile], [Timestamp], [DataArea]"

    stgTable = KeyFam & "_kf_stg"
    DropTable stgTable

    If rowCount > 0 Then
        DbExecute "SELECT * INTO " & InB(stgTable) & " FROM " & InB(KeyFam)
    End If

    DropTable KeyFam

    DbExecute "CREATE TABLE " & InB(KeyFam) & " (" & colDefs & ")"

    Select Case CurrentDB.DBType
        Case accdb, mdb
            DbExecute "CREATE UNIQUE INDEX [PrimaryIndex] ON " & InB(KeyFam) & " (" & idxCols & ") WITH PRIMARY"
        Case Sqlexpress
            DbExecute "CREATE UNIQUE INDEX [PrimaryIndex] ON " & InB(KeyFam) & " (" & idxCols & ")"
    End Select

    If rowCount > 0 Then
        DbExecute "INSERT INTO " & InB(KeyFam) & " (" & insCols & ") SELECT " & selCols & " FROM " & InB(stgTable)
        DropTable stgTable
    End If

    On Error Resume Next
    If Not CurrentDB.DBCat Is Nothing Then
        CurrentDB.DBCat.Tables.Refresh
    End If
    On Error GoTo errHandler

    CloseDB

    CurrentDB.DimensionsIsLoaded = False
    CurrentDB.KeynamesIsLoaded = False
    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions

    MsgBox "Dimension '" & DimName & "' (Length: " & DimLen & ") was successfully added to Key Family '" & KeyFam & "'.", vbInformation, "NADABAS"
    KeyFamilyAddDimension = True
    Exit Function

errHandler:
    errNum = err.Number
    errDesc = err.Description
    CloseDB
    ShowNadabasError "KeyFamilyAddDimension", errNum, errDesc
    KeyFamilyAddDimension = False
End Function

Public Function KeyFamilyRemoveDimension(KeyFam As String, DimName As String) As Boolean
    Dim keyf As clsKeyName
    Dim field As clsFieldNames
    Dim remainingDims As Collection
    Dim idxCols As String
    Dim colDefs As String
    Dim insCols As String
    Dim selCols As String
    Dim ssql As String
    Dim dups As Long
    Dim rowCount As Long
    Dim valType As Integer
    Dim valColDef As String
    Dim stgTable As String
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo errHandler
    KeyFamilyRemoveDimension = False

    OpenDb

    Set remainingDims = New Collection
    Set keyf = CurrentDB.GetKeyName(KeyFam)
    valType = 2

    If Not keyf Is Nothing Then
        For Each field In keyf.TableDefinition
            If UCase(field.name) = "VALUE" Then
                valType = field.SQLType
                Exit For
            End If
            If UCase(field.name) <> UCase(DimName) Then
                remainingDims.Add field
            End If
        Next field
    End If

    If remainingDims.count < 2 Then
        CloseDB
        MsgBox "A Key Family must contain at least 2 dimensions." & vbCrLf & _
               "Cannot remove dimension '" & DimName & "'.", vbCritical, "NADABAS"
        Exit Function
    End If

    If MsgBox("Are you sure you want to permanently remove dimension '" & DimName & "' from Key Family '" & KeyFam & "'?", _
              vbYesNo + vbQuestion, "NADABAS") = vbNo Then
        CloseDB
        Exit Function
    End If

    idxCols = ""
    For Each field In remainingDims
        If idxCols <> "" Then idxCols = idxCols & ", "
        idxCols = idxCols & InB(field.name)
    Next field

    rowCount = 0
    If CreateCursor("SELECT COUNT(*) AS Cnt FROM " & InB(KeyFam)) Then
        If Not CursorEoF Then
            rowCount = CLng(GetColumn("Cnt"))
        End If
        CloseCursor
    End If

    If rowCount > 0 Then
        ssql = "SELECT COUNT(*) AS DupCount FROM (SELECT " & idxCols & ", COUNT(*) AS Cnt FROM " & InB(KeyFam) & _
               " GROUP BY " & idxCols & " HAVING COUNT(*) > 1) AS DupQuery"
        If CreateCursor(ssql) Then
            If Not CursorEoF Then
                dups = CLng(GetColumn("DupCount"))
            End If
            CloseCursor
        End If

        If dups > 0 Then
            CloseDB
            MsgBox "Cannot remove dimension '" & DimName & "'." & vbCrLf & vbCrLf & _
                   "Removing it would produce " & dups & " duplicate key combination(s) among existing records.", _
                   vbCritical, "NADABAS"
            Exit Function
        End If
    End If

    Select Case valType
        Case ADOX.DataTypeEnum.adSingle:
            valColDef = "SINGLE"
        Case ADOX.DataTypeEnum.adDouble:
            valColDef = "DOUBLE"
        Case ADOX.DataTypeEnum.adVarWChar:
            valColDef = IIf(CurrentDB.DBType = Sqlexpress, "NVARCHAR(255)", "TEXT(255)")
        Case Else
            valColDef = "DOUBLE"
    End Select

    colDefs = ""
    idxCols = ""
    insCols = ""
    selCols = ""

    For Each field In remainingDims
        If colDefs <> "" Then colDefs = colDefs & ", "
        If idxCols <> "" Then idxCols = idxCols & ", "
        If insCols <> "" Then insCols = insCols & ", "
        If selCols <> "" Then selCols = selCols & ", "

        Select Case CurrentDB.DBType
            Case accdb, mdb
                colDefs = colDefs & InB(field.name) & " TEXT(" & field.Length & ") NOT NULL"
            Case Sqlexpress
                colDefs = colDefs & InB(field.name) & " NVARCHAR(" & field.Length & ") NOT NULL"
        End Select

        idxCols = idxCols & InB(field.name)
        insCols = insCols & InB(field.name)
        selCols = selCols & InB(field.name)
    Next field

    Select Case CurrentDB.DBType
        Case accdb, mdb
            colDefs = colDefs & ", [Value] " & valColDef & _
                                ", [Comment] MEMO" & _
                                ", [Formula] TEXT(255)" & _
                                ", [Username] TEXT(50)" & _
                                ", [ExcelFile] TEXT(255)" & _
                                ", [Timestamp] DATETIME" & _
                                ", [DataArea] TEXT(50)"
        Case Sqlexpress
            colDefs = colDefs & ", [Value] " & valColDef & _
                                ", [Comment] NVARCHAR(MAX)" & _
                                ", [Formula] NVARCHAR(255)" & _
                                ", [Username] NVARCHAR(50)" & _
                                ", [ExcelFile] NVARCHAR(255)" & _
                                ", [Timestamp] DATETIME" & _
                                ", [DataArea] NVARCHAR(50)"
    End Select

    insCols = insCols & ", [Value], [Comment], [Formula], [Username], [ExcelFile], [Timestamp], [DataArea]"
    selCols = selCols & ", [Value], [Comment], [Formula], [Username], [ExcelFile], [Timestamp], [DataArea]"

    stgTable = KeyFam & "_kf_stg"
    DropTable stgTable

    If rowCount > 0 Then
        DbExecute "SELECT * INTO " & InB(stgTable) & " FROM " & InB(KeyFam)
    End If

    DropTable KeyFam

    DbExecute "CREATE TABLE " & InB(KeyFam) & " (" & colDefs & ")"

    Select Case CurrentDB.DBType
        Case accdb, mdb
            DbExecute "CREATE UNIQUE INDEX [PrimaryIndex] ON " & InB(KeyFam) & " (" & idxCols & ") WITH PRIMARY"
        Case Sqlexpress
            DbExecute "CREATE UNIQUE INDEX [PrimaryIndex] ON " & InB(KeyFam) & " (" & idxCols & ")"
    End Select

    If rowCount > 0 Then
        DbExecute "INSERT INTO " & InB(KeyFam) & " (" & insCols & ") SELECT " & selCols & " FROM " & InB(stgTable)
        DropTable stgTable
    End If

    On Error Resume Next
    If Not CurrentDB.DBCat Is Nothing Then
        CurrentDB.DBCat.Tables.Refresh
    End If
    On Error GoTo errHandler

    CloseDB

    CurrentDB.DimensionsIsLoaded = False
    CurrentDB.KeynamesIsLoaded = False
    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions

    MsgBox "Dimension '" & DimName & "' was successfully removed from Key Family '" & KeyFam & "'.", vbInformation, "NADABAS"
    KeyFamilyRemoveDimension = True
    Exit Function

errHandler:
    errNum = err.Number
    errDesc = err.Description
    CloseDB
    ShowNadabasError "KeyFamilyRemoveDimension", errNum, errDesc
    KeyFamilyRemoveDimension = False
End Function

Public Function KeyFamilyReorderIndex(KeyFam As String, DimOrder As Collection) As Boolean
    Dim keyf As clsKeyName
    Dim field As clsFieldNames
    Dim dimMap As Collection
    Dim idxCols As String
    Dim colDefs As String
    Dim insCols As String
    Dim selCols As String
    Dim v As Variant
    Dim rowCount As Long
    Dim valType As Integer
    Dim valColDef As String
    Dim stgTable As String
    Dim dName As String
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo errHandler
    KeyFamilyReorderIndex = False

    If DimOrder.count < 2 Then Exit Function

    OpenDb

    Set dimMap = New Collection
    Set keyf = CurrentDB.GetKeyName(KeyFam)
    valType = 2

    If Not keyf Is Nothing Then
        For Each field In keyf.TableDefinition
            If UCase(field.name) = "VALUE" Then
                valType = field.SQLType
                Exit For
            End If
            On Error Resume Next
            dimMap.Add field, UCase(field.name)
            On Error GoTo 0
        Next field
    End If

    If dimMap.count = 0 Then
        CloseDB
        Exit Function
    End If

    rowCount = 0
    If CreateCursor("SELECT COUNT(*) AS Cnt FROM " & InB(KeyFam)) Then
        If Not CursorEoF Then
            rowCount = CLng(GetColumn("Cnt"))
        End If
        CloseCursor
    End If

    Select Case valType
        Case ADOX.DataTypeEnum.adSingle:
            valColDef = "SINGLE"
        Case ADOX.DataTypeEnum.adDouble:
            valColDef = "DOUBLE"
        Case ADOX.DataTypeEnum.adVarWChar:
            valColDef = IIf(CurrentDB.DBType = Sqlexpress, "NVARCHAR(255)", "TEXT(255)")
        Case Else
            valColDef = "DOUBLE"
    End Select

    colDefs = ""
    idxCols = ""
    insCols = ""
    selCols = ""

    For Each v In DimOrder
        dName = CStr(v)
        Set field = dimMap(UCase(dName))

        If colDefs <> "" Then colDefs = colDefs & ", "
        If idxCols <> "" Then idxCols = idxCols & ", "
        If insCols <> "" Then insCols = insCols & ", "
        If selCols <> "" Then selCols = selCols & ", "

        Select Case CurrentDB.DBType
            Case accdb, mdb
                colDefs = colDefs & InB(field.name) & " TEXT(" & field.Length & ") NOT NULL"
            Case Sqlexpress
                colDefs = colDefs & InB(field.name) & " NVARCHAR(" & field.Length & ") NOT NULL"
        End Select

        idxCols = idxCols & InB(field.name)
        insCols = insCols & InB(field.name)
        selCols = selCols & InB(field.name)
    Next v

    Select Case CurrentDB.DBType
        Case accdb, mdb
            colDefs = colDefs & ", [Value] " & valColDef & _
                                ", [Comment] MEMO" & _
                                ", [Formula] TEXT(255)" & _
                                ", [Username] TEXT(50)" & _
                                ", [ExcelFile] TEXT(255)" & _
                                ", [Timestamp] DATETIME" & _
                                ", [DataArea] TEXT(50)"
        Case Sqlexpress
            colDefs = colDefs & ", [Value] " & valColDef & _
                                ", [Comment] NVARCHAR(MAX)" & _
                                ", [Formula] NVARCHAR(255)" & _
                                ", [Username] NVARCHAR(50)" & _
                                ", [ExcelFile] NVARCHAR(255)" & _
                                ", [Timestamp] DATETIME" & _
                                ", [DataArea] NVARCHAR(50)"
    End Select

    insCols = insCols & ", [Value], [Comment], [Formula], [Username], [ExcelFile], [Timestamp], [DataArea]"
    selCols = selCols & ", [Value], [Comment], [Formula], [Username], [ExcelFile], [Timestamp], [DataArea]"

    stgTable = KeyFam & "_kf_stg"
    DropTable stgTable

    If rowCount > 0 Then
        DbExecute "SELECT * INTO " & InB(stgTable) & " FROM " & InB(KeyFam)
    End If

    DropTable KeyFam

    DbExecute "CREATE TABLE " & InB(KeyFam) & " (" & colDefs & ")"

    Select Case CurrentDB.DBType
        Case accdb, mdb
            DbExecute "CREATE UNIQUE INDEX [PrimaryIndex] ON " & InB(KeyFam) & " (" & idxCols & ") WITH PRIMARY"
        Case Sqlexpress
            DbExecute "CREATE UNIQUE INDEX [PrimaryIndex] ON " & InB(KeyFam) & " (" & idxCols & ")"
    End Select

    If rowCount > 0 Then
        DbExecute "INSERT INTO " & InB(KeyFam) & " (" & insCols & ") SELECT " & selCols & " FROM " & InB(stgTable)
        DropTable stgTable
    End If

    On Error Resume Next
    If Not CurrentDB.DBCat Is Nothing Then
        CurrentDB.DBCat.Tables.Refresh
    End If
    On Error GoTo errHandler

    CloseDB

    CurrentDB.DimensionsIsLoaded = False
    CurrentDB.KeynamesIsLoaded = False
    CurrentDB.LoadKeyNames
    CurrentDB.LoadDimensions

    KeyFamilyReorderIndex = True
    Exit Function

errHandler:
    errNum = err.Number
    errDesc = err.Description
    CloseDB
    ShowNadabasError "KeyFamilyReorderIndex", errNum, errDesc
    KeyFamilyReorderIndex = False
End Function
