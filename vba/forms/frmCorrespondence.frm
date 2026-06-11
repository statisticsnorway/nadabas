Attribute VB_Name = "frmCorrespondence"
Attribute VB_Base = "0{B37B18B7-2B9B-4958-9F77-BDFEF1828653}{657E8540-4C4E-48D2-837D-EF74B940B739}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Returncode As Integer
Public CorrSelected As clsCorrespondence
Private WithEvents mcTree As clsTreeView
Attribute mcTree.VB_VarHelpID = -1

Public Sub Initialize(AsAdministrator As Boolean)
Dim corr As clsCorrespondence
Dim n As Integer

 
    
    lbLabels.Clear
    lbLabels.ColumnCount = 2
    lbLabels.ColumnWidths = "99;"
    lbLabels.AddItem
    lbLabels.Column(0, 0) = Me.lblSource.Caption
    lbLabels.Column(1, 0) = Me.lblTarget.Caption
 
    lbCorrespondences.Clear
    lbCorrespondences.ColumnCount = 2
    lbCorrespondences.ColumnWidths = "99;"

    n = 0
    For Each corr In CurrentDB.Correspondences
        lbCorrespondences.AddItem
        lbCorrespondences.Column(0, n) = corr.Sourceclass
        lbCorrespondences.Column(1, n) = corr.TargetClass
        n = n + 1
    Next corr
 
    If n > 0 Then
       lbCorrespondences.ListIndex = 0
'       Filltreeview CurrentDB.Correspondences(1)
    End If
    
    cmdAdd.Visible = AsAdministrator
    cmdAdd.Enabled = AsAdministrator
    
    cmdEdit.Visible = AsAdministrator
    cmdEdit.Enabled = AsAdministrator & lbCorrespondences.ListIndex > -1
    
    cmdDelete.Visible = AsAdministrator
    cmdDelete.Enabled = AsAdministrator & lbCorrespondences.ListIndex > -1
 

End Sub
Private Sub Filltreeview(corr As clsCorrespondence)
Dim oldTargetCode As String
Dim oldSourceCode As String

Dim CItem As clsCorrItem
Dim n As Integer
Dim CurrParent As String
Dim Title As String
Dim Class As clsClassification
Dim cRoot As clsNode
Dim cNode As clsNode
  
    If Not mcTree Is Nothing Then
        mcTree.TerminateTree
        Set mcTree = Nothing
    End If
    Set mcTree = New clsTreeView
    With mcTree
        Set .TreeControl = Me.frTreeControl
        Call .NodesClear
      '  .AppName = Me.AppName
        .RootButton = True
        .EnableLabelEdit = False
        .FullWidth = False
        .Indentation = 15
        .NodeHeight = 10.5
        .ShowLines = True
        .ShowExpanders = True
    End With

    n = 0
    mcTree.NodesClear
    oldSourceCode = ""

    For Each CItem In corr.Items
       n = n + 1
       If CItem.TargetCode <> oldTargetCode Then
          CurrParent = "P" & n
         Set Class = CurrentDB.Classifications(corr.TargetClass)
          Title = CItem.TargetCode & " " & Class.GetClassItemTitle(CItem.TargetCode)
          Set cRoot = mcTree.AddRoot(sKey:=CurrParent, vCaption:=Title)
          oldTargetCode = CItem.TargetCode
          cRoot.Expanded = True
       End If
         Set Class = CurrentDB.Classifications(corr.Sourceclass)
         Title = CItem.SourceCode & " " & Class.GetClassItemTitle(CItem.SourceCode)
         cRoot.AddChild "C" & n, Title
    Next CItem
    mcTree.Refresh



End Sub


Private Sub cmdAdd_Click()
  Returncode = 1
  Me.Hide
End Sub

Private Sub cmdDelete_Click()
Dim corr As clsCorrespondence
    Set corr = CurrentDB.Correspondences(lbCorrespondences.ListIndex + 1)
     If MsgBox(GetMsg2("M075", corr.Sourceclass, corr.TargetClass), vbYesNo) = vbNo Then Exit Sub   'confirm to delete Correspondence between %1 and %2
      corr.DeleteInDB
      CurrentDB.LoadCorrespondences     ' load existing correspondences if any
      Initialize True
End Sub




Private Sub cmdEdit_Click()
  If lbCorrespondences.ListIndex < 0 Then Exit Sub
  Set CorrSelected = CurrentDB.Correspondences(lbCorrespondences.ListIndex + 1)
  Returncode = 2
  Me.Hide
End Sub


Private Sub cmdExport_Click()
'
' export all Correspondences to a workbook
'
' the workbook will contain 1 sheet for each source classification
'
' firs to columns are the code and title of the source, then follow 2 columns for each target with code and title
'
 
     Returncode = 3
     Me.Hide
End Sub

Private Sub cmdOK_Click()
  Returncode = 0
  Me.Hide
End Sub

Private Sub lbCorrespondences_Click()
Dim corr As clsCorrespondence

    If lbCorrespondences.ListIndex < 0 Then Exit Sub
       Set corr = CurrentDB.Correspondences(lbCorrespondences.ListIndex + 1)
    Filltreeview corr
End Sub


Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub

