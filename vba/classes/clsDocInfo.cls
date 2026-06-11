Attribute VB_Name = "clsDocInfo"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public name As String
Public path As String
Public Level As String
Public DGroup As String
Public Workbook As String
 


Public Sub AddToDB()
    OpenDb

    CreateCursor "Select * from documents where 1 <> 1 "
    CursorAddNew
    PutColumn "Name", name
    PutColumn "path", path
    PutColumn "Level", Level
    PutColumn "DGroup", DGroup
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
