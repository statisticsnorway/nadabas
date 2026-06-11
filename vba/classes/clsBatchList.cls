Attribute VB_Name = "clsBatchList"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public name As String
Public Entries As Collection

Private Sub Class_Initialize()
    Set Entries = New Collection    ' a collection of BatchElements
End Sub

Public Sub savetoBatchlist()
Dim BE As clsBatchListEntry
    
    OpenDb
    '
    ' get rid of any old list
    '
    DbExecute "delete  from batchlist where listname = " & InQ(name)
    
    CreateCursor "select * from batchlist where 1 <> 1"
    For Each BE In Entries
        CursorAddNew
        PutColumn "Listname", name
        PutColumn "ItemNo", BE.Itemno
        PutColumn "WorkbookName", BE.ElementName
        CursorUpdate
    Next BE
    CloseCursor
 
    CloseDB
End Sub

Public Sub savetoBatchlist2()
Dim BE As clsBatchListEntry
    
    OpenDb
    '
    ' get rid of any old list
    '
    DbExecute "delete  from batch2list where listname = " & InQ(name)
    
    CreateCursor "select * from batch2list where 1 <> 1"
    For Each BE In Entries
        CursorAddNew
        PutColumn "Listname", name
        PutColumn "ItemNo", BE.Itemno
        PutColumn "Batchname", BE.ElementName
        CursorUpdate
    Next BE
    CloseCursor
 
    CloseDB
End Sub
