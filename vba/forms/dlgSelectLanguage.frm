Attribute VB_Name = "dlgSelectLanguage"
Attribute VB_Base = "0{81912DC9-9378-4DCB-873A-2BB0AF49B4C4}{CF8CE368-BD82-44FA-BEE7-23EFEC085FF6}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Dim lancode As Integer

Private Sub OptionButton1_Click()
lancode = 3
End Sub

Private Sub OptionButton2_Click()
  lancode = 4
End Sub

Private Sub OptionButton3_Click()
    lancode = 5
End Sub


Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
'    Translateform Me    ' translate all labels etc.
    OptionButton1.value = True
   OptionButton1_Click ' click it to set lancode
End Sub

Private Sub cmdCancel_Click()
   Me.Hide
End Sub

Private Sub cmdOK_Click()
 SetLanguageSetting lancode
 Me.Hide
End Sub
