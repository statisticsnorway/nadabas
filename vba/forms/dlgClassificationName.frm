Attribute VB_Name = "dlgClassificationName"
Attribute VB_Base = "0{8FCE1EA3-B906-4033-99BE-9A72C6DCE215}{CB41434D-075B-4DDF-B085-253666C1CD97}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Public cancel As Boolean
 
Private Sub cmdCancel_Click()
   cancel = True
   Me.Hide
End Sub

Private Sub cmdOK_Click()
   If Trim(txtClassName) = "" Then
       MsgBox GetMsg("M052"), vbOKOnly       'Please enter a classname
       Exit Sub
   End If
   cancel = False
   Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
 
