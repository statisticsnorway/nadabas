VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBConstant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' used to save constants from DBLinks and record if they are used (during Test)
'

Public DBConstant As String
Public Used As Boolean

Private Sub Class_Initialize()
    Used = False
    DBConstant = ""
End Sub
