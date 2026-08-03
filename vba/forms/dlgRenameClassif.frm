VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgRenameClassif
   Caption         =   "Rename Classification"
   ClientHeight    =   2310
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4590
   OleObjectBlob   =   "dlgRenameClassif.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgRenameClassif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NewName As String

Private Sub cmdCancel_Click()
NewName = ""
  Me.Hide
End Sub

Private Sub cmdOK_Click()
NewName = Me.txtNewName.Text
  Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
