VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWorkBookInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WorkBookName As String
Public Title As String
Public GroupName As String
Public LastGet As String
Public LastPut As String
Public RelPath As String  ' without basepath (whats in the DB)
Public path As String    ' in normal form, basepath has been added
Public NeedsUpdate As Boolean
Public HasMissing As Boolean
Public ReservedBy As String
Public ReservedDate As String

Public Isdirty As Long
Public Status As String

Public DontCopy As Boolean      ' used during transfer to mark files to be skipped

Public FileExists As Boolean    ' used by test all xls only (not from database


Public Sub SaveInDB()

    OpenDb
    CreateCursor "select * from workbooks where Workbookname = " & InQ(WorkBookName)
    If CursorEoF Then
       CursorAddNew
    Else
       CursorEdit
    End If
    PutColumn "Workbookname", WorkBookName
    PutColumn "Title", Title
    PutColumn "GroupName", GroupName
    PutColumn "Path", RelPath
    PutColumnNull "Status", Status
    PutColumnNull "ReservedBy", ReservedBy
    PutColumnNull "ReservedDate", ReservedDate
    PutColumn "isdirty", 0
    CursorUpdate
    CloseCursor
    CloseDB
End Sub
'

Public Sub DeleteFromDB()
   OpenDb
   DbExecute "Delete from workbooks where WorkBookName = " & InQ(WorkBookName)
   CloseDB
End Sub

Public Sub FreeOrReserveinDB()
        OpenDb
        CreateCursor "select * from workbooks where Workbookname = " & InQ(WorkBookName)
    If Not CursorEoF Then
        CursorEdit
        PutColumnNull "ReservedBy", ReservedBy
        PutColumnNull "ReservedDate", ReservedDate
        PutColumnNull "Status", Status
        CursorUpdate
    End If

    CloseCursor
    CloseDB
End Sub

Private Sub Class_Initialize()
   WorkBookName = ""
   Title = ""
   LastGet = ""
   LastPut = ""
   path = ""
   NeedsUpdate = False
   HasMissing = False
   Isdirty = 0
   Status = ""
   FileExists = False

End Sub
