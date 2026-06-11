Attribute VB_Name = "dlgSelectDatabase"
Attribute VB_Base = "0{8B84B002-2EA3-4C08-AF6A-EE33B5763126}{1A2C38C2-46E1-41D6-8397-F4E3FCCAE7DC}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public selectedbase As Integer
Public selectedName As String

Private Sub cmdCancel_Click()
    selectedbase = -1                ' no database selected
    selectedName = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
     selectedbase = lbBases.ListIndex
     selectedName = lbBases.Text
     If selectedbase < 0 Then
         MsgBox GetMsg("M062"), vbOKOnly  'Please select one
        Exit Sub
     End If
     Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
