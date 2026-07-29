VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPermission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WorkBookName As String
Public user As String

Public Property Get key() As String
      key = WorkBookName & "_" & user
End Property
Public Sub SaveToDB()
    OpenDb
    If Not DBTableExists("Permissions") Then
        CreateTablePermissions
    End If

    CreateCursor "Select * from Permissions where 1 <> 1"

    CursorAddNew
    PutColumn "Workbookname", DropFileType(WorkBookName)
    PutColumn "User", user
    CursorUpdate
    CloseCursor
    CloseDB
End Sub

Public Sub DeleteFromDB()
    OpenDb
    DbExecute "Delete from Permissions where workbookname = " & _
            InQ(WorkBookName) & "and [user] = " & _
            InQ(user)
    CloseDB
End Sub
