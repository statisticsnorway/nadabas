VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDimClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Dimensionname As String
Public classname As String

Public Sub DeleteFromDB()
    OpenDb
    DbExecute "Delete from DimensionClass where DimensionName = " & InQ(Dimensionname)
    CloseDB
End Sub

Public Sub SaveToDB()
    OpenDb
    CreateCursor "select * from DimensionClass where DimensionName = " & InQ(Dimensionname)
    If Not CursorEoF Then
       CursorEdit
    Else
       CursorAddNew
    End If
    PutColumn "DimensionName", Dimensionname
    PutColumn "ClassName", classname
    CursorUpdate
    CloseCursor
    CloseDB
End Sub
