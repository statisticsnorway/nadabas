VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgBackUp2
   Caption         =   "Backup"
   ClientHeight    =   5448
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9930.001
   OleObjectBlob   =   "dlgBackUp2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgBackUp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public BackupSelected As String


Private Sub cmdOK_Click()
    If lbFiles.ListIndex = -1 Then
        MsgBox "Please select a backup folder", vbExclamation
        Exit Sub
    End If

    Dim folderName As String
    folderName = lbFiles.value  ' just the folder name

    ' ? Store just the folder name
    BackupSelected = folderName

    Me.Hide
    Unload Me
End Sub


Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
