Attribute VB_Name = "dlgPutLog"
Attribute VB_Base = "0{8E005477-E49B-4D1B-99FA-E050407B7CEB}{2EBDAFB1-0B9D-418C-BFE5-A2F8B0499289}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Public LogYesNo As Boolean

Private Sub cmdDetail_Click()
  lstBoxDetails.Visible = True
  Me.Height = 630
End Sub

Private Sub cmdNo_Click()
 Me.Hide
 LogYesNo = False
End Sub

Private Sub cmdYes_Click()
  Me.Hide
  LogYesNo = True
End Sub


Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
