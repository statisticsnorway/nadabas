VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgAbout
   Caption         =   "NADABAS"
   ClientHeight    =   11985
   ClientLeft      =   -15
   ClientTop       =   195
   ClientWidth     =   5235
   OleObjectBlob   =   "dlgAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
