Attribute VB_Name = "frmCorrespondanceEdit"
Attribute VB_Base = "0{87F3B137-CC83-4040-843F-6E3208372172}{145566E2-798E-4631-BFE2-EC47C8E0F400}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Returncode As Integer

Dim NewCorr As clsCorrespondence ' the Correspondance under construction
Dim Source As clsClassification  ' only those no in newcorr.items
Dim Sourceclass As clsClassification
Dim TargetClass As clsClassification
Private WithEvents mcTree As clsTreeView
Attribute mcTree.VB_VarHelpID = -1


Public Sub Initialize(corr As clsCorrespondence)
Dim Coitem As clsCorrItem
Dim CClass As clsClassification
Dim ClItem As clsClassItem
'
' correspondences and classifications are loaded
'
'  copy from existing correspondance




  Set Source = New clsClassification
  Source.classname = corr.Sourceclass
  Set Sourceclass = CurrentDB.Classifications(corr.Sourceclass)
  Set TargetClass = CurrentDB.Classifications(corr.TargetClass)
  
  For Each ClItem In Sourceclass.Items
      Source.Items.Add ClItem, ClItem.code
  Next ClItem
  
  Set NewCorr = New clsCorrespondence
  
  NewCorr.Sourceclass = corr.Sourceclass
  NewCorr.TargetClass = corr.TargetClass
  
  For Each Coitem In corr.Items
      NewCorr.AddItemSorted Coitem
      Source.Items.Remove Coitem.SourceCode       ' remove from source
  Next Coitem

  FillTargetTree
  FillSourceList

 End Sub


Private Sub FillTargetTree()
Dim ClItem As clsClassItem
Dim Coitem As clsCorrItem
Dim Class As clsClassification
Dim Title As String
Dim CurrParent As String
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
        .RootButton = True
        .EnableLabelEdit = False
        .FullWidth = False
        .Indentation = 15
        .NodeHeight = 10.5
        .ShowLines = True
        .ShowExpanders = True
    End With
    
   mcTree.NodesClear
   For Each ClItem In TargetClass.Items
      Set cRoot = mcTree.AddRoot(sKey:="T" & ClItem.code, vCaption:=ClItem.code & " " & ClItem.Title)
      Set cRoot.tag = ClItem
      cRoot.Expanded = True
   Next ClItem
   
   For Each Coitem In NewCorr.Items
          Title = Coitem.SourceCode & " " & Sourceclass.GetClassItemTitle(Coitem.SourceCode)
          CurrParent = "T" & Coitem.TargetCode
          Set cRoot = mcTree.Nodes(CurrParent)
          Set cNode = cRoot.AddChild("S" & Coitem.SourceCode, Title)
          Set cNode.tag = Coitem
   Next Coitem
   mcTree.Refresh

End Sub


Private Sub FillSourceList()

Dim ClItem As clsClassItem

    lbSourceCodes.Clear
    
    For Each ClItem In Source.Items
       lbSourceCodes.AddItem ClItem.code & " " & ClItem.Title
    Next ClItem

End Sub

Private Sub cmdAddItem_Click()
Dim Tnode As clsNode
Dim n As Integer
Dim Coitem As clsCorrItem
Dim clTarget As clsClassItem
Dim ClItem As clsClassItem
    If mcTree.ActiveNode Is Nothing Then
       MsgBox GetMsg("M072"), vbOKOnly     ' Please select a target
       Exit Sub
    End If
    Set Tnode = mcTree.ActiveNode
        
    If (Mid(Tnode.key, 1, 1) <> "T") Then
         MsgBox GetMsg("M072"), vbOKOnly     ' Please select a target
       Exit Sub
    End If
    
  Set clTarget = Tnode.tag
  For n = 0 To lbSourceCodes.ListCount - 1
       If lbSourceCodes.Selected(n) = True Then
           Set ClItem = Source.Items(n + 1)
           Set Coitem = New clsCorrItem
           Coitem.SourceCode = ClItem.code
           Coitem.TargetCode = clTarget.code
           NewCorr.AddItemSorted Coitem
        End If
  Next n
        
     For n = lbSourceCodes.ListCount - 1 To 0 Step -1
       If lbSourceCodes.Selected(n) = True Then
          lbSourceCodes.RemoveItem (n)
           Source.Items.Remove (n + 1)
       End If
   Next n
   FillTargetTree
   FillSourceList
        
End Sub


Private Sub cmdEditInExcel_Click()
   Returncode = 1
   Me.Hide
End Sub

Private Sub cmdRemoveitem_Click()
Dim Tnode As clsNode
Dim n As Integer
Dim Coitem As clsCorrItem
Dim ClItem As clsClassItem
    
    If mcTree.ActiveNode Is Nothing Then
       MsgBox GetMsg("M073"), vbOKOnly    'Please select souce node to remove", vbOKOnly
       Exit Sub
    End If
    Set Tnode = mcTree.ActiveNode
        
    If (Mid(Tnode.key, 1, 1) = "T") Then
        MsgBox GetMsg("M074"), vbOKOnly    'Selected item must be a source node
       Exit Sub
    End If
    
    Set Coitem = Tnode.tag
    NewCorr.DeleteItem Coitem
    
    Source.AddItemSorted Sourceclass.Items(Coitem.SourceCode)
 
    FillTargetTree
    FillSourceList
    
    
End Sub



Private Sub cmdCancel_Click()
    Returncode = 0
    Me.Hide
End Sub

Private Sub cmdSave_Click()
    
    NewCorr.SaveInDB
    CurrentDB.LoadCorrespondences     ' reload correspondences if any
    Returncode = 0
    Me.Hide

End Sub

Private Sub UserForm_Initialize()
    Translateform Me    ' translate all labels etc.
End Sub
