Attribute VB_Name = "dlgBackupType"
Attribute VB_Base = "0{734B926E-9894-49CC-AECA-B45F4460DD5A}{32B2FF19-3C77-4C98-A84F-9D9BE9BA6FE6}"
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
  cancel = False
  Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
