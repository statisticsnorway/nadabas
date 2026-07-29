VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSatelliteImport
   Caption         =   "Select workbooks to import"
   ClientHeight    =   9645.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   14910
   OleObjectBlob   =   "frmSatelliteImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSatelliteImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Returncode As Boolean
Public WBsToOpen As Collection
Private tags As Collection

Public Sub Initialize(WBs As Collection)
Dim WBinfo As clsWorkBookInfo
Dim v As Variant
Dim maxfilenameLen As Integer


    maxfilenameLen = CurrentDB.MaxFileNameLength

    If maxfilenameLen < 17 Then
       ListBox1.ColumnWidths = 80
       lbTitles.ColumnWidths = 80
    Else
        ListBox1.ColumnWidths = maxfilenameLen * 5
        lbTitles.ColumnWidths = maxfilenameLen * 5
   End If

    lbTitles.AddItem Me.lblFileName.Caption
    lbTitles.ListIndex = lbTitles.ListCount - 1
    lbTitles.Column(1) = Me.lblTitle.Caption
    lbTitles.ListIndex = -1


     ListBox1.Clear
     Set tags = New Collection
     For Each WBinfo In WBs
            With ListBox1
                .AddItem
                .List(.ListCount - 1, 0) = WBinfo.WorkBookName
                .List(.ListCount - 1, 1) = WBinfo.Title
            End With
            tags.Add WBinfo
     Next WBinfo
End Sub



Private Sub cmdCancel_Click()
       Returncode = False
        Me.Hide
End Sub

Private Sub cmdImport_Click()
Dim n As Integer

    Set WBsToOpen = New Collection

    With ListBox1
        For n = 0 To .ListCount - 1
            If .Selected(n) = True Then
                WBsToOpen.Add tags(n + 1)
            End If
        Next n
    End With
    If WBsToOpen.count = 0 Then
       MsgBox GetMsg("M153"), vbInformation
       Exit Sub
    End If
    Returncode = True
    Me.Hide

End Sub

Private Sub UserForm_Initialize()
    DropClose Me               ' get rid of Close button on frame
    Translateform Me    ' translate all labels etc.
End Sub
