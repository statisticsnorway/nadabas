Attribute VB_Name = "dlgLog"
Attribute VB_Base = "0{466BAF34-E26D-4E79-8C03-14C9F255172A}{B92DB6BD-ECC3-4733-A8BD-E379D6C1D474}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Public LogYesNo As Boolean


Private Sub cmdNo_Click()
 Me.Hide
 LogYesNo = False
End Sub

Private Sub cmdOK_Click()
  Me.Hide
  LogYesNo = True
End Sub

Private Sub cmdYes_Click()
  Me.Hide
  LogYesNo = True
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
