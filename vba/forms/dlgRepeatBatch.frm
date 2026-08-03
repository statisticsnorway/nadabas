VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgRepeatBatch
   Caption         =   "Repeated Batch"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   OleObjectBlob   =   "dlgRepeatBatch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgRepeatBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public MaxRepeatNo As Long
Public cancel As Boolean

Private Sub cmdCancel_Click()
   cancel = True
   Me.Hide
End Sub

Private Sub cmdOK_Click()
On Error GoTo someerror
   cancel = False
   MaxRepeatNo = txtMax
   If MaxRepeatNo < 1 Or MaxRepeatNo > 50 Then GoTo someerror
   Me.Hide
   Exit Sub
someerror:
   MsgBox GetMsg("M061"), vbOKOnly   'Repeat must be a number between 1 and 50

End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
