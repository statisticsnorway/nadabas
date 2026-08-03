VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgRunBatch
   Caption         =   "Nadabas Run Batch Update"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   OleObjectBlob   =   "dlgRunBatch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgRunBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
