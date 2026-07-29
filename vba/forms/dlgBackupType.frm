VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgBackupType
   Caption         =   "Select type of backup"
   ClientHeight    =   3936
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "dlgBackupType.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgBackupType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cancel As Boolean


Private Sub cmdCancel_Click()
  cancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
  cancel = False
  Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
