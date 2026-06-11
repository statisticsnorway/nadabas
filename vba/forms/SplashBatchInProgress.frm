Attribute VB_Name = "SplashBatchInProgress"
Attribute VB_Base = "0{20E3C154-146E-4739-9C99-6038D29941B9}{A40654E3-6134-43A4-B701-F5CA0D896B56}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub CommandButton1_Click()
  If MsgBox("Cancel update ", vbYesNo) = vbYes Then
     BatchRunInProgress = False
     Me.Textcancel.Caption = Me.lblCanncelIn.Caption
     Me.CommandButton1.Enabled = False
     DoEvents
   End If
End Sub

 Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
