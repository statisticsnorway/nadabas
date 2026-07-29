VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgPutLog
   Caption         =   "Save data conflict"
   ClientHeight    =   11760
   ClientLeft      =   120
   ClientTop       =   660
   ClientWidth     =   12030
   OleObjectBlob   =   "dlgPutLog.frx":0000
End
Attribute VB_Name = "dlgPutLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public LogYesNo As Boolean

Private Sub cmdDetail_Click()
  lstBoxDetails.Visible = True
  Me.Height = 630
End Sub

Private Sub cmdNo_Click()
 Me.Hide
 LogYesNo = False
End Sub

Private Sub cmdYes_Click()
  Me.Hide
  LogYesNo = True
End Sub


Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
