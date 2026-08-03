VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgFindMarkedItems
   Caption         =   "Marked Cells"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "dlgFindMarkedItems.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgFindMarkedItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
  Me.Hide
  Unload Me
End Sub

Private Sub ListBox1_Click()
Dim s As String
Dim s1 As String
Dim x As Range
  On Error Resume Next
  s1 = ListBox1.value
  s = ListBox1.Column(1)
  Range(s1).Worksheet.Activate
  Range(s1).Worksheet.Range(s).Select
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
