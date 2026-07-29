VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgVersion
   Caption         =   "Version control"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "dlgVersion.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOK_Click()
 Usersettings.NoVersionControl = cbDontShow.value
 Usersettings.SaveVersionControl
 Me.Hide
End Sub


Private Sub UserForm_Initialize()

  DropClose Me               ' get rid of Close button on frame
'
   Translateform Me    ' translate all labels etc.
   Me.cbDontShow.value = False
End Sub
