Attribute VB_Name = "cmdBaseFolder"
Option Explicit
Option Private Module
'
' Anything relted to the Basepath goes here
'

Public Sub SetBaseFolder()
' *******************
' called from Ribbon
' *******************
    If CurrentDB.BasefolderInfo Is Nothing Then
       CurrentDB.LoadBasePath
    End If

     Load frmBaseFolder

     frmBaseFolder.Initialize CurrentDB.BasefolderInfo

     frmBaseFolder.Show vbModal
     If frmBaseFolder.cancel = False Then
        OpenDb
        If frmBaseFolder.CheckBox1.value = True Then
          CurrentDB.BasefolderInfo.basePath = "!"
        Else
          CurrentDB.BasefolderInfo.basePath = DropBackSlash(frmBaseFolder.txtFolder)
        End If
        saveBasePath
        Set CurrentDB.BasefolderInfo = Nothing
        CurrentDB.LoadBasePath           ' in order to get it rigth relative to DB
        RepairWorkbookPathsForBaseFolder
        CurrentDB.LoadWorkbookInfo       ' refresh paths resolved against the new base folder
        CloseDB

        MsgBox GetMsg("M033") & vbCrLf & CurrentDB.BasefolderInfo.basePath, vbOKOnly 'Base folder changed to

     End If
     Unload frmBaseFolder
End Sub

Public Function RepairWorkbookPathsForBaseFolder() As Long
'
' Check every registered workbook for a portable ! path after Set base folder.
' Absolute paths already below the selected base folder are converted directly.
' Other missing/non-portable paths are repaired only when the workbook name
' identifies one folder below the selected base folder.
'
    Dim AmbiguousNames As Object
    Dim DirectPaths As Object
    Dim FoundFolders As Object
    Dim RepairIDs As Object
    Dim RepairNames As Object
    Dim NameKey As Variant
    Dim portablePath As String
    Dim StoredPath As String
    Dim WorkbookID As Long
    Dim WorkbookName As String

    If Not DBTableExists("Workbooks") Then Exit Function
    If CurrentDB.BasefolderInfo Is Nothing Then Exit Function
    If Len(Trim$(CurrentDB.BasefolderInfo.basePath)) = 0 Then Exit Function

    Set RepairIDs = CreateObject("Scripting.Dictionary")
    RepairIDs.CompareMode = vbTextCompare
    Set RepairNames = CreateObject("Scripting.Dictionary")
    RepairNames.CompareMode = vbTextCompare
    Set DirectPaths = CreateObject("Scripting.Dictionary")
    DirectPaths.CompareMode = vbTextCompare
    Set FoundFolders = CreateObject("Scripting.Dictionary")
    FoundFolders.CompareMode = vbTextCompare
    Set AmbiguousNames = CreateObject("Scripting.Dictionary")
    AmbiguousNames.CompareMode = vbTextCompare

    CreateCursor "SELECT [WorkbookID], [WorkBookName], [Path] FROM [Workbooks]"

    Do While Not CursorEoF
        WorkbookID = CLng(GetColumn("WorkbookID"))
        WorkbookName = CStr(GetColumn("WorkBookName"))
        StoredPath = CStr(GetColumnNull("Path"))

        If Not IsPortableWorkbookPath(StoredPath) Then
            RepairIDs(WorkbookName) = WorkbookID
            RepairNames(WorkbookName) = WorkbookName

            If Len(Trim$(StoredPath)) > 0 Then
                If TestBasePath(StoredPath) Then
                    DirectPaths(WorkbookName) = _
                        DropBasePath(DropBackSlash(StoredPath))
                End If
            End If
        End If
        CursorMoveNext
    Loop
    CloseCursor

    If RepairIDs.count = 0 Then Exit Function

    For Each NameKey In DirectPaths.Keys
        RepairWorkbookPathsForBaseFolder = _
            RepairWorkbookPathsForBaseFolder + DbExecute( _
                "UPDATE [Workbooks] SET [Path] = " & _
                InQ(CStr(DirectPaths(NameKey))) & _
                " WHERE [WorkbookID] = " & CStr(RepairIDs(NameKey)) _
            )
        RepairIDs.Remove NameKey
        RepairNames.Remove NameKey
    Next NameKey

    If RepairIDs.count > 0 Then
        ScanFolderForMissingWorkbooks CurrentDB.BasefolderInfo.basePath, _
                                      RepairIDs, FoundFolders, AmbiguousNames
    End If

    For Each NameKey In FoundFolders.Keys
        If Not AmbiguousNames.Exists(CStr(NameKey)) Then
            portablePath = DropBasePath(CStr(FoundFolders(NameKey)))
            RepairWorkbookPathsForBaseFolder = _
                RepairWorkbookPathsForBaseFolder + DbExecute( _
                "UPDATE [Workbooks] SET [Path] = " & InQ(portablePath) & _
                " WHERE [WorkbookID] = " & CStr(RepairIDs(NameKey)) _
            )
        End If
    Next NameKey

    For Each NameKey In RepairNames.Keys
        If Not FoundFolders.Exists(CStr(NameKey)) Then
            AddToSQLLog "Set base folder could not locate workbook: " & _
                        CStr(RepairNames(NameKey))
        ElseIf AmbiguousNames.Exists(CStr(NameKey)) Then
            AddToSQLLog "Set base folder found workbook in multiple folders: " & _
                        CStr(RepairNames(NameKey))
        End If
    Next NameKey

End Function

Private Function IsPortableWorkbookPath(ByVal StoredPath As String) As Boolean
    Dim trimmedPath As String

    trimmedPath = Trim$(StoredPath)
    If StrComp(StoredPath, trimmedPath, vbBinaryCompare) <> 0 Then Exit Function

    If trimmedPath = "!" Then
        IsPortableWorkbookPath = True
    ElseIf Len(trimmedPath) > 2 Then
        IsPortableWorkbookPath = Left$(trimmedPath, 2) = "!\" Or _
                                 Left$(trimmedPath, 2) = "!/"
    End If

End Function

Private Sub ScanFolderForMissingWorkbooks( _
    ByVal FolderPath As String, _
    ByVal RepairIDs As Object, _
    ByVal FoundFolders As Object, _
    ByVal AmbiguousNames As Object _
)
    Dim extension As String
    Dim file As Object
    Dim files As Object
    Dim subfolder As Object
    Dim subfolders As Object
    Dim WorkbookName As String

    On Error GoTo FolderUnavailable

    Set files = GetFilesInFolder(FolderPath)
    For Each file In files
        extension = LCase$(GetFileExtension(CStr(file.name)))
        Select Case extension
            Case "xls", "xlsx", "xlsm", "xlsb"
                WorkbookName = DropFileType(CStr(file.name))
                If RepairIDs.Exists(WorkbookName) Then
                    If FoundFolders.Exists(WorkbookName) Then
                        If StrComp(CStr(FoundFolders(WorkbookName)), FolderPath, _
                                   vbTextCompare) <> 0 Then
                            AmbiguousNames(WorkbookName) = True
                        End If
                    Else
                        FoundFolders(WorkbookName) = FolderPath
                    End If
                End If
        End Select
    Next file

    Set subfolders = GetSubFolders(FolderPath)
    For Each subfolder In subfolders
        ScanFolderForMissingWorkbooks CStr(subfolder.path), RepairIDs, _
                                      FoundFolders, AmbiguousNames
    Next subfolder

    DoEvents
    Exit Sub

FolderUnavailable:
    AddToSQLLog "Set base folder could not scan folder: " & FolderPath & _
                vbCrLf & CStr(err.Number) & ":" & err.Description

End Sub



Public Function getBasepath() As String
    If CurrentDB.BasefolderInfo Is Nothing Then
       CurrentDB.LoadBasePath
    End If
    getBasepath = CurrentDB.BasefolderInfo.basePath
End Function
Public Sub SetBasePathForAccDB()
     Set CurrentDB.BasefolderInfo = New clsBasePath
     CurrentDB.BasefolderInfo.basePath = "!"
     CurrentDB.BasefolderInfo.BasePathIsDB = True
     saveBasePath
     Set CurrentDB.BasefolderInfo = Nothing        ' must be loaded correctly (! to be replaced by path)

End Sub
Public Sub saveBasePath()

     If Not DBTableExists("Basepath") Then
        CreateTableBasePath
      End If
      CurrentDB.BasefolderInfo.saveBasePathToDB

End Sub


Public Function TestBasePath(InPath As String) As Boolean
Dim x As Long
Dim Y As Long
Dim bpa As String
    If CurrentDB.BasefolderInfo Is Nothing Then
       CurrentDB.LoadBasePath
    End If
    TestBasePath = False
    If NadabasIsSleeping Then Exit Function
    TestBasePath = True
    If CurrentDB.BasefolderInfo.basePath = "" Then Exit Function
    If Len(InPath) > Len(CurrentDB.BasefolderInfo.basePath) Then
       bpa = CurrentDB.BasefolderInfo.basePath & "\"
    Else
      bpa = CurrentDB.BasefolderInfo.basePath
    End If
    x = Len(bpa)
    Y = Len(InPath)
    If Y >= x And UCase(Mid(InPath, 1, x)) = UCase(bpa) Then
       Exit Function
    End If
    TestBasePath = False
End Function
