Attribute VB_Name = "dlgVersion"
Attribute VB_Base = "0{79842FCA-5D84-47EE-829C-6EFFF9F05EF5}{4AB992CD-71CF-423B-9246-13D5E732B6B6}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Sub cmdOK_Click()
 Usersettings.NoVersionControl = cbDontShow.value
 Usersettings.SaveVersionControl
 Me.Hide
End Sub


Private Sub UserForm_Initialize()
  
  DropClose Me               ' get rid of Close button on frame
'
   Translateform Me    ' translate all labels etc.
   Me.cbDontShow.value = False
End Sub
