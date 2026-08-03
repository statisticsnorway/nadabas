VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBackUpFolder
   Caption         =   "Select backup folder for "
   ClientHeight    =   3708
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9420.001
   OleObjectBlob   =   "frmBackUpFolder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBackUpFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public cancel As Boolean

Public Sub Initialize(BackUpPath As clsBackUpPath)
    txtFolder.Text = GetPath2(BackUpPath.BackUpPath)
    txtTemplate.Text = GetFilename(BackUpPath.BackUpPath)
    Caption = Me.Caption & " " & CurrentDB.DbDisplayName
End Sub

Private Sub cmdFind_Click()

Dim s As String
    s = BrowseFolder(Me.lblSelectFolder.Caption)
    If s <> "" Then
        txtFolder = s
    End If
End Sub
Private Sub cmdCancel_Click()
    cancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
   Dim bp As String
      bp = DropBackSlash(Me.txtFolder)
      If Len(bp) >= Len(getBasepath) Then
        If getBasepath = Mid(bp, 1, Len(getBasepath)) Then
           MsgBox GetMsg("M050"), vbOKOnly   'Backuppath must not be within basepath
           Exit Sub
        End If
      End If
    If txtTemplate = "" Then
      MsgBox GetMsg("M051"), vbOKOnly         'Please enter template for folder
      Exit Sub
    End If
    cancel = False
    Me.Hide
End Sub


Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
