VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDocInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public name As String
Public path As String
Public Level As String
Public DGroup As String
Public WorkbookID As Long
Public Workbook As String



Public Sub AddToDB()
    OpenDb

    CreateCursor "Select * from documents where 1 <> 1 "
    CursorAddNew
    PutColumn "Name", name
    PutColumn "path", path
    PutColumn "Level", Level
    PutColumn "DGroup", DGroup
    If WorkbookID = 0 And Len(Workbook) <> 0 Then
        WorkbookID = GetOrCreateWorkbookID(Workbook)
    End If
    PutColumnNull "WorkbookID", IIf(WorkbookID = 0, "", WorkbookID)
    PutColumn "Workbook", Workbook
    CursorUpdate
    CloseCursor
    CloseDB
 End Sub

Public Sub DeleteFromDB()
    OpenDb
    DbExecute "Delete  from documents where  Name = " & InQ(name)
    CloseDB
End Sub
