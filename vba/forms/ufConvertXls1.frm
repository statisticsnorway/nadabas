Attribute VB_Name = "ufConvertXls1"
Attribute VB_Base = "0{478C9557-6D19-42B1-A46E-29BDD4D31968}{6B6447B8-56E8-4F85-8300-02114E924CCB}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Public cancel As Boolean

Private Sub CommandButton1_Click()
   
   cancel = False
   Me.Hide
End Sub

Private Sub CommandButton2_Click()
  cancel = True
  Me.Hide
End Sub

Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
