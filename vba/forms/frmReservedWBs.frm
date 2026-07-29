VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReservedWBs
   Caption         =   "Workbooks that are reserved of transferred"
   ClientHeight    =   9675.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   16035
   OleObjectBlob   =   "frmReservedWBs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReservedWBs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub Initialize()
Dim SI As clsWorkBookInfo
Dim n As Long
Dim subPath As String
Dim fullname As String

    OpenDb

    CurrentDB.LoadKeyNames
    CurrentDB.LoadWorkbookInfo    ' make sure we got as fresh copy
    CloseDB

    lbLabels.Clear
    lbLabels.ColumnCount = 5
    lbLabels.ColumnWidths = "120;120;90;90;"
    lbLabels.AddItem
    lbLabels.Column(0, 0) = Me.lblGroupName.Caption
    lbLabels.Column(1, 0) = Me.lblWorkbook.Caption
    lbLabels.Column(2, 0) = Me.lblStatus.Caption
    lbLabels.Column(3, 0) = Me.lblBy.Caption
    lbLabels.Column(4, 0) = Me.lbldate.Caption


    lbSheets.Clear

    lbSheets.ColumnCount = 5
    lbSheets.ColumnWidths = "120;120;90;90;"

    n = 0
    For Each SI In CurrentDB.WorkBooks
       If SI.ReservedBy <> "" Then
            lbSheets.AddItem
            lbSheets.Column(0, n) = SI.GroupName
            lbSheets.Column(1, n) = SI.WorkBookName
            lbSheets.Column(2, n) = SI.Status
            lbSheets.Column(3, n) = SI.ReservedBy
            lbSheets.Column(4, n) = SI.ReservedDate
            n = n + 1
       End If
    Next SI

    If lbSheets.ListCount = 0 Then
        lbSheets.AddItem
        lbSheets.Column(0, n) = Me.lblNoReserved.Caption
    End If


End Sub



Private Sub cmdFinish_Click()
  Me.Hide
End Sub


Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
