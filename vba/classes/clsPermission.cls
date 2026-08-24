VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPermission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WorkbookID As Long
Public WorkbookName As String
Public user As String

Public Property Get key() As String
      key = CStr(WorkbookID) & "_" & user
End Property
Public Sub SaveToDB()
    OpenDb
    If Not DBTableExists("Permissions") Then
        CreateTablePermissions
    End If

    CreateCursor "Select * from Permissions where 1 <> 1"

    CursorAddNew
    If WorkbookID = 0 Then
        WorkbookID = GetOrCreateWorkbookID(WorkBookName)
    End If
    PutColumn "WorkbookID", WorkbookID
    PutColumn "Workbookname", WorkBookName
    PutColumn "User", user
    CursorUpdate
    CloseCursor
    CloseDB
End Sub

Public Sub DeleteFromDB()
    OpenDb
    If WorkbookID = 0 Then
        WorkbookID = GetWorkbookID(WorkBookName)
    End If
    DbExecute "Delete from Permissions where WorkbookID = " & _
            CStr(WorkbookID) & " and [user] = " & _
            InQ(user)
    CloseDB
End Sub
