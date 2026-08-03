VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufConvertXls1
   Caption         =   "Convert to xlsb"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8250.001
   OleObjectBlob   =   "ufConvertXls1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufConvertXls1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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
