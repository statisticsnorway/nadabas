Attribute VB_Name = "dlgRepeatBatch"
Attribute VB_Base = "0{9CE4E350-B866-4E1E-B953-9621819C18AE}{217FC288-05D6-4E5F-A3D4-D286CA046094}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public MaxRepeatNo As Long
Public cancel As Boolean

Private Sub cmdCancel_Click()
   cancel = True
   Me.Hide
End Sub

Private Sub cmdOK_Click()
On Error GoTo someerror
   cancel = False
   MaxRepeatNo = txtMax
   If MaxRepeatNo < 1 Or MaxRepeatNo > 50 Then GoTo someerror
   Me.Hide
   Exit Sub
someerror:
   MsgBox GetMsg("M061"), vbOKOnly   'Repeat must be a number between 1 and 50

End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
