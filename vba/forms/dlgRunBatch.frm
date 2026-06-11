Attribute VB_Name = "dlgRunBatch"
Attribute VB_Base = "0{7AD3A2D0-730C-41AD-9FED-EA41FE37B444}{D0569117-8390-4E50-A92A-7B2222D14706}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Public cancel As Boolean
 


Private Sub CancelButton_Click()
   cancel = True
   Me.Hide
End Sub

Private Sub OkButton_Click()
   cancel = False
   Me.Hide
End Sub

 
Private Sub UserForm_Initialize()
    Me.cbCloseWB.value = True
    Me.cbSilentMode.value = True
    Me.cbFormulaIgnore.value = False
    Me.cbDropbox.Visible = False
    Me.cbDropbox.value = False
  
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
    
End Sub


