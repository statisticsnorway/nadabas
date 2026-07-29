VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClassDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public classname As String
Public Description As String

Public Sub SaveToDB()
    OpenDb
    'ClassificationDescriptions
    CreateCursor "Select * from ClassificationDescriptions where 1 <> 1"
    CursorAddNew
    PutColumn "ClassName", classname
    PutColumn "Description", Description
    CursorUpdate
    CloseCursor
    CloseDB
End Sub
Public Sub DeleteFromDB()
    OpenDb
    DbExecute "Delete from ClassificationDescriptions where classname = " & InQ(classname)
    CloseDB
End Sub
