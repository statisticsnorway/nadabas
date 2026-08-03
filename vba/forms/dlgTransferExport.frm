VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgTransferExport
   Caption         =   "Transfer workbook to sattelite"
   ClientHeight    =   3036
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4590
   OleObjectBlob   =   "dlgTransferExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgTransferExport"
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
    Returncode = True
    Me.Hide
End Sub


Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
