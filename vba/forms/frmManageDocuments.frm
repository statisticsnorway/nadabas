Attribute VB_Name = "frmManageDocuments"
Attribute VB_Base = "0{DB323175-B6B4-4081-AE0B-201BC041ED23}{4C45880A-529C-4EFF-B553-6BED953E9406}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

 

Public Sub Initialize()
Dim Di As clsDocInfo
Dim n As Long
    CurrentDB.LoadDocuments
   
    lbLabels.Clear
    lbLabels.ColumnCount = 5
    lbLabels.ColumnWidths = "120;60;80;120;"
    lbLabels.AddItem
    lbLabels.Column(0, 0) = Me.lblDocument.Caption
    lbLabels.Column(1, 0) = Me.lblLevel.Caption
    lbLabels.Column(2, 0) = Me.lblGroup.Caption
    lbLabels.Column(3, 0) = Me.lblWorkbook.Caption
    lbLabels.Column(4, 0) = Me.lblPath.Caption
   
    lbSheets.Clear
  
    lbSheets.ColumnCount = 5
    lbSheets.ColumnWidths = "120;60;80;120;"

    n = 0
    For Each Di In CurrentDB.Documents
        lbSheets.AddItem
        lbSheets.Column(0, n) = Di.name
        lbSheets.Column(1, n) = Di.Level
        lbSheets.Column(2, n) = Di.DGroup
        lbSheets.Column(3, n) = Di.Workbook

        If fsFileExists(Di.path & "\" & Di.name) Then
          lbSheets.Column(4, n) = Di.path
        Else
          lbSheets.Column(4, n) = Me.lblNotFound.Caption & " " & Di.path
        End If
        n = n + 1
    Next Di
End Sub


Private Sub cmdAdd_Click()
  RegisterDocument
  Initialize
End Sub

Private Sub cmdDelete_Click()
Dim x As Long
Dim Di As clsDocInfo
    x = lbSheets.ListIndex + 1
    Set Di = CurrentDB.Documents(x)
    Di.DeleteFromDB
    Initialize
End Sub


Private Sub cmdOK_Click()
  Me.Hide
End Sub


Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
