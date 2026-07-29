VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgRenameBatch
   Caption         =   "RenameBatchfile"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "dlgRenameBatch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgRenameBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public listname As String

Private Sub cmdCancel_Click()
    listname = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
      listname = Me.TextBox1.Text
'
' test not in use
'
     If CurrentDB.BatchlistnameExist(listname) Then
         MsgBox GetMsg("M060"), vbCritical      'Name in use
        Exit Sub
     End If
     Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
