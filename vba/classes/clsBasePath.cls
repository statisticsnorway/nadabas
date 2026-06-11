Attribute VB_Name = "clsBasePath"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public basePath As String


Public BasePathIsDB As Boolean

Public Sub LoadBasePathFromDB()

    basePath = ""
    If NadabasIsSleeping Then Exit Sub
    If Not DBTableExists("Basepath") Then Exit Sub
    OpenDb
    CreateCursor "Select Basepath  from Basepath"
    If Not CursorEoF Then
        basePath = GetColumn("Basepath")
    End If
    CloseCursor
    CloseDB
    BasePathIsDB = False
    If basePath = "!" Then
       basePath = GetPath(CurrentDB.DBFullName)
       BasePathIsDB = True
    End If
    basePath = DropBackSlash(basePath)
    
End Sub

 Public Sub saveBasePathToDB()
 '
' delete any old basepath
'
 

    OpenDb
    DbExecute "Delete  from BasePath"
   
'
' add the one and only row needed
'
   DbExecute "Insert into BAsePath values(" & InQ(basePath) & ")"
    
    CloseDB

End Sub
 
