VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBatchDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public name As String
Public Description As String

Public Sub SaveDescriptionToDB()
    OpenDb
    DbExecute "delete  from batchdescription where listname = " & InQ(name)
    CreateCursor "select * from batchdescription where 1 <> 1"
    CursorAddNew
    PutColumn "Listname", name
    PutColumn "Description", Description
    CursorUpdate
    CloseCursor
    CloseDB
End Sub

Public Sub SaveDescription2ToDB()
    OpenDb
    DbExecute "delete  from batch2description where listname = " & InQ(name)
    CreateCursor "select * from batch2description where 1 <> 1"
    CursorAddNew
    PutColumn "Listname", name
    PutColumn "Description", Description
    CursorUpdate
    CloseCursor
    CloseDB
End Sub
