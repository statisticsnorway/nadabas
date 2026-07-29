VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgSelectLanguage
   Caption         =   "Select Language"
   ClientHeight    =   3024
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3390
   OleObjectBlob   =   "dlgSelectLanguage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgSelectLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
