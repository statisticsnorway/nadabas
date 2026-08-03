VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWorkBookStatus
   Caption         =   "All workbooks"
   ClientHeight    =   9225.001
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15345
   OleObjectBlob   =   "frmWorkBookStatus.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWorkBookStatus"
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
    lbLabels.ColumnCount = 9
    lbLabels.ColumnWidths = "120;120;90;90;3;3;30;30;"
    lbLabels.AddItem
    lbLabels.Column(0, 0) = Me.lblGroupName.Caption
    lbLabels.Column(1, 0) = Me.lblWorkbook.Caption
    lbLabels.Column(2, 0) = Me.LblLastPut.Caption
    lbLabels.Column(3, 0) = Me.lblLastGet.Caption
    lbLabels.Column(4, 0) = " "   ' not usec anymore
    lbLabels.Column(5, 0) = " "   'not used anymore
    lblMiss.Visible = False
    lbLabels.Column(6, 0) = Me.lblRes.Caption
    lbLabels.Column(7, 0) = Me.lblDirty.Caption
    lbLabels.Column(8, 0) = Me.lblPath.Caption

    lbSheets.Clear

    lbSheets.ColumnCount = 9
    lbSheets.ColumnWidths = "120;120;90;90;3;3;30;30;"
    lbSheets.AddItem

    n = 0
    For Each SI In CurrentDB.WorkBooks
       lbSheets.AddItem
       lbSheets.Column(0, n) = SI.GroupName
       lbSheets.Column(1, n) = SI.WorkBookName
       lbSheets.Column(2, n) = SI.LastPut
       lbSheets.Column(3, n) = SI.LastGet

    ' no col 4,5
       If SI.ReservedBy <> "" Then

          lbSheets.Column(6, n) = Mid(SI.Status, 1, 1)
       End If
       If SI.Isdirty <> 0 Then
          lbSheets.Column(7, n) = "X"
       End If

       subPath = Replace(SI.path, getBasepath, "!", 1, 1, vbTextCompare)
       fullname = GetFullWorkbookName(SI.path & "\" & SI.WorkBookName)
       If MultipleFiles Then
          lbSheets.Column(8, n) = subPath & " " & Me.lblMultiple.Caption
       Else
       If fullname <> "" Then
         lbSheets.Column(8, n) = subPath
       Else
          lbSheets.Column(8, n) = Me.lblNotFound.Caption & " " & subPath
        End If
        End If
       n = n + 1
    Next SI

End Sub


Private Sub cdmDetails_Click()
Dim x As Long
Dim SI As clsWorkBookInfo

    x = lbSheets.ListIndex
    If x < 0 Then Exit Sub
    Set SI = CurrentDB.WorkBooks(x + 1)
    Load dlgSheetDetails
    dlgSheetDetails.LoadData SI
    dlgSheetDetails.Show vbModal
    Unload dlgSheetDetails
End Sub

Private Sub cmdFinish_Click()
  Me.Hide
End Sub

Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
