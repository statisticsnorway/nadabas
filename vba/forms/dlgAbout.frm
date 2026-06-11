Attribute VB_Name = "dlgAbout"
Attribute VB_Base = "0{C967E320-CEFD-456B-BF2F-6A0C98B07E66}{E286CCBB-6AF9-4BAB-A71B-4ABE67D76EFA}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Private Sub cmdErrors_Click()
    OpenNadabasError
End Sub

Private Sub cmdOK_Click()
  Me.Hide
End Sub

Private Sub cmdVideo_Click()
    OpenVideo
End Sub

Private Sub cmdWEB_Click()
     OpenNadabasWeb
End Sub


Private Sub Label15_Click()

End Sub

Private Sub UserForm_Initialize()
  
  DropClose Me               ' get rid of Close button on frame
'
   Translateform Me    ' translate all labels etc.
End Sub

