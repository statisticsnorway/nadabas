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

Public WorkbookID As Long
Public WorkbookName As String
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
    If WorkbookID = 0 Then
        WorkbookID = GetOrCreateWorkbookID(WorkBookName)
    End If

    CreateCursor "select * from workbooks where WorkbookID = " & CStr(WorkbookID)
    If CursorEoF Then
       CursorAddNew
    Else
       CursorEdit
    End If
    PutColumn "WorkbookID", WorkbookID
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
   DbExecute "Delete from workbooks where WorkbookID = " & CStr(WorkbookID)
   CloseDB
End Sub

Public Sub FreeOrReserveinDB()
        OpenDb
        CreateCursor "select * from workbooks where WorkbookID = " & CStr(WorkbookID)
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
   WorkbookID = 0
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
