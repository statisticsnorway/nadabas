VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgExportWB
   Caption         =   "Export workbook"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "dlgExportWB.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgExportWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Returncode As Boolean

Private Sub cmdCancel_Click()
   Returncode = False
   Me.Hide
End Sub

Private Sub cmdOK_Click()
    If txtUsername.Text = "" Then
       MsgBox GetMsg("M152"), vbOKOnly
       Exit Sub
    End If
    Returncode = True
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
