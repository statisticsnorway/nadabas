VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgSelectExcelFile
   Caption         =   "Select File"
   ClientHeight    =   5472
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   11025
   OleObjectBlob   =   "dlgSelectExcelFile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgSelectExcelFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FileSelected As String
Public cancel As Boolean

Public Sub Initialize(basePath As String, FName As String)

Dim n As Integer
Dim files As Object
Dim file As Variant

    Set files = GetFilesInFolder(basePath)

    lbLabels.Clear
    lbLabels.ColumnCount = 3
    lbLabels.ColumnWidths = "120;90;"
    lbLabels.AddItem
    lbLabels.Column(0, 0) = Me.lblFile.Caption
    lbLabels.Column(1, 0) = Me.lblCreated.Caption
    lbLabels.Column(2, 0) = Me.lblLastChange.Caption

    lbSheets.Clear

    lbSheets.ColumnCount = 3
    lbSheets.ColumnWidths = "120;90;"

    lbSheets.AddItem
    n = 0
    For Each file In files
    If DropFileType(file.name) = FName Then
       lbSheets.AddItem
      lbSheets.Column(0, n) = file.name
      lbSheets.Column(1, n) = file.DateCreated
      lbSheets.Column(2, n) = file.DateLastModified

      n = n + 1
    End If
     Next file
End Sub





Private Sub cmdCancel_Click()
         cancel = True
         FileSelected = ""
         Me.Hide
End Sub

Private Sub cmdOK_Click()
         If lbSheets.ListIndex < 0 Then Exit Sub
         cancel = False
         FileSelected = lbSheets.Column(0, lbSheets.ListIndex)
         Me.Hide
End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
