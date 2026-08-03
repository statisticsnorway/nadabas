VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgAccToSQL
   Caption         =   "Convert .mdb to SQL"
   ClientHeight    =   3210
   ClientLeft      =   30
   ClientTop       =   345
   ClientWidth     =   4410
   OleObjectBlob   =   "dlgAccToSQL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgAccToSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cancel As Boolean

Private Sub cmdCancel_Click()
   cancel = True
   Me.Hide
End Sub

Private Sub cmdOK_Click()
   cancel = False
   Me.Hide
End Sub


Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
