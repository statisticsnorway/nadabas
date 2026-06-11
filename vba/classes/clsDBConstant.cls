Attribute VB_Name = "clsDBConstant"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'
' used to save constants from DBLinks and record if they are used (during Test)
'

Public DBConstant As String
Public Used As Boolean

Private Sub Class_Initialize()
    Used = False
    DBConstant = ""
End Sub
