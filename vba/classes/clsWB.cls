VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
