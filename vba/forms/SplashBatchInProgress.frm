VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SplashBatchInProgress
   Caption         =   "Batch Update in Progress"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "SplashBatchInProgress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "SplashBatchInProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
