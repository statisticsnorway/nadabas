VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKeyFamilyMigrationButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ActionButton As MSForms.CommandButton
Attribute ActionButton.VB_VarHelpID = -1
Private Owner As frmKeyFamily
Private ChangeType As String

Public Sub Initialize(Button As MSForms.CommandButton, _
                      FormOwner As frmKeyFamily, _
                      SchemaChangeType As String)
    Set ActionButton = Button
    Set Owner = FormOwner
    ChangeType = SchemaChangeType
End Sub

Private Sub ActionButton_Click()
    Owner.SchemaChangeButtonClick ChangeType
End Sub
