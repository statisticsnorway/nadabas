VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgCreateNewAcc
   Caption         =   "New Access DB"
   ClientHeight    =   2760
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   13470
   OleObjectBlob   =   "dlgCreateNewAcc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgCreateNewAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public cancel As Boolean


Private Sub cmdFind_Click()

Dim s As String
    s = BrowseFolder("Select folder")
    If s <> "" Then
        txtFolder = s
    End If
End Sub
Private Sub cmdCancel_Click()
    cancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()

    If txtFolder = "" Then
      MsgBox GetMsg("M150"), vbOKOnly     'Please enter template for folder
      Exit Sub
    End If
    If txtFilename = "" Then
      MsgBox GetMsg("M151"), vbOKOnly 'Please enter fielname
      Exit Sub
    End If
    cancel = False
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
