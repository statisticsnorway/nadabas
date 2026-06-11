Attribute VB_Name = "clsWB"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public name As String
Public bPut As Boolean
Public bGet As Boolean
Public Number As Long   ' used by missing only
'
Private Sub Class_Initialize()
   bPut = False
   bGet = False
   Number = 0
End Sub
 
 
