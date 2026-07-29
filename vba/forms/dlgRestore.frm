VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgRestore
   Caption         =   "Restore from backup"
   ClientHeight    =   5748
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7260
   OleObjectBlob   =   "dlgRestore.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cancel As Boolean
Public BackupSelected As String
Dim BackUpBase As String

Public Sub Initialize(BackUpPath As String)
    ' Populate ListBox1 with subfolder names under BackUpPath

    Dim fld As Object
    Dim subFolder As Object

    Set fld = InterfaceFileScripting.GetSubFolders(BackUpPath)

    ListBox1.Clear

    For Each subFolder In fld
        ' ? Only add the folder name, not full path
        Me.ListBox1.AddItem subFolder.name
    Next subFolder
End Sub



Private Sub cmdOK_Click()
  If ListBox1.ListIndex = -1 Then
     MsgBox GetMsg("M062"), vbOKOnly  'Please select one
     Exit Sub
  End If
  BackupSelected = BackUpBase & ListBox1.Text
  cancel = False
  Me.Hide
End Sub


Private Sub cmdCancel_Click()
 cancel = True
 Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
