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
Public ApplyToAll As Boolean
Private WithEvents cmdYesAll As MSForms.CommandButton
Private WithEvents cmdNoAll As MSForms.CommandButton

Private Sub cmdYesAll_Click()
    ApplyToAll = True
    cmdYes_Click
End Sub

Private Sub cmdNoAll_Click()
    ApplyToAll = True
    cmdNo_Click
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        LogYesNo = False
        ApplyToAll = False
        Me.Hide
    End If
End Sub

Private Sub cmdDetail_Click()
  lstBoxDetails.Visible = True
  Me.Height = 630
  lstBoxDetails.Height = Me.InsideHeight - lstBoxDetails.Top - 12
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
    LogYesNo = False
    ApplyToAll = False
    Set cmdYesAll = Me.Controls.Add("Forms.CommandButton.1", "cmdYesAll", True)
    Set cmdNoAll = Me.Controls.Add("Forms.CommandButton.1", "cmdNoAll", True)
    With cmdYesAll
        .Caption = "Yes to all"
        .Left = cmdYes.Left
        .Top = cmdYes.Top + cmdYes.Height + 6
        .Width = cmdYes.Width
        .Height = cmdYes.Height
        .TabIndex = cmdNo.TabIndex + 1
    End With
    With cmdNoAll
        .Caption = "No to all"
        .Left = cmdNo.Left
        .Top = cmdYesAll.Top
        .Width = cmdNo.Width
        .Height = cmdNo.Height
        .TabIndex = cmdYesAll.TabIndex + 1
    End With
    ' Leave room for the extra row when the details list is expanded.
    lstBoxDetails.Top = cmdYesAll.Top + cmdYesAll.Height + 12
    cmdNo.Cancel = True
    cmdNo.Default = True
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
