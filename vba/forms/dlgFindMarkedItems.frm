Attribute VB_Name = "dlgFindMarkedItems"
Attribute VB_Base = "0{EDCE3BEE-87D3-4746-BD5A-024E4910B997}{E2C38F7E-4E5B-43D8-8F4E-FC902281B65B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Private Sub CommandButton1_Click()
  Me.Hide
  Unload Me
End Sub

Private Sub ListBox1_Click()
Dim s As String
Dim s1 As String
Dim x As Range
  On Error Resume Next
  s1 = ListBox1.value
  s = ListBox1.Column(1)
  Range(s1).Worksheet.Activate
  Range(s1).Worksheet.Range(s).Select
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
