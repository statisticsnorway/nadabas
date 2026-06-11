Attribute VB_Name = "clsClassDescription"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
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
