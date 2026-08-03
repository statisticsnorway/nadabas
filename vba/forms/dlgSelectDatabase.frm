VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgSelectDatabase
   Caption         =   "Select Database"
   ClientHeight    =   4452
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   12585
   OleObjectBlob   =   "dlgSelectDatabase.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgSelectDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public selectedbase As Integer
Public selectedName As String

Private Sub cmdCancel_Click()
    selectedbase = -1                ' no database selected
    selectedName = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
     selectedbase = lbBases.ListIndex
     selectedName = lbBases.Text
     If selectedbase < 0 Then
         MsgBox GetMsg("M062"), vbOKOnly  'Please select one
        Exit Sub
     End If
     Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
