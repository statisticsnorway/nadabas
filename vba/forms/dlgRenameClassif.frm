Attribute VB_Name = "dlgRenameClassif"
Attribute VB_Base = "0{0A89F96E-7D90-487B-8EE3-524D2A183E51}{DEF40DDF-B884-4479-A785-9862C776BABD}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public NewName As String

Private Sub cmdCancel_Click()
NewName = ""
  Me.Hide
End Sub

Private Sub cmdOK_Click()
NewName = Me.txtNewName.Text
  Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
