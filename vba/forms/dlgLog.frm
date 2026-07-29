VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgLog
   Caption         =   "Log"
   ClientHeight    =   8865.001
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615.001
   OleObjectBlob   =   "dlgLog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public LogYesNo As Boolean


Private Sub cmdNo_Click()
 Me.Hide
 LogYesNo = False
End Sub

Private Sub cmdOK_Click()
  Me.Hide
  LogYesNo = True
End Sub

Private Sub cmdYes_Click()
  Me.Hide
  LogYesNo = True
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
